<#
.SYNOPSIS
一条命令跑完 Office-helper 的全部回归：构建 + 4 组 .NET 测试 + 全部 web 单测。

.DESCRIPTION
之前每个验证脚本各管一段（宿主加载、面板、端到端），发版前要手工敲四五条命令，
漏跑一段不会有任何提示。这个脚本把它们串成一条链路，任一步非零退出即整体失败，
最后给出「通过 N，失败 M，跳过 S」的汇总，并逐条列出失败项。

步骤：
  1. dotnet build ChatSheet.sln --configuration <Configuration>
  2. ChatSheet.ToolTests.exe --agent-tool-tests      （流式、能力、工具、请求形状）
  3. ChatSheet.ToolTests.exe --reliability-tests    （可靠性回归）
  4. ChatSheet.ToolTests.exe --project-audit-tests  （需要真实 Excel，-SkipHost 可跳过）
  5. ChatSheet.ToolTests.exe --codex-config-tests   （本地 fixture，不联网）
  6. tests\web\*.test.mjs 逐个 node 跑

每一步的完整输出都会落到 artifacts\verify-all\
  下，控制台只保留实时输出；失败时直接指向对应日志。

.PARAMETER Configuration
构建配置，默认 Release。Debug 也能跑，但 Release 是发布前的验收配置。

.PARAMETER SkipHost
跳过全部需要真实宿主（Excel）的步骤，即 --project-audit-tests。
没有装 Excel 的机器、或 Excel 正被别的任务占用时用这个开关，
其余步骤照常执行，结果里记为「跳过」而不是「通过」。

.PARAMETER IncludePanelUi
额外跑需要桌面会话和 WebView2 的面板验证（PaneHarness --panel-security / --workbuddy-ui）。
默认不跑：这两步依赖真实桌面、WebView2 运行时和本机 WorkBuddy 授权状态，
放进默认链路会让一次提交前的快速验证变得不可预期。

.EXAMPLE
.\scripts\verify-all.ps1
完整验证，包含需要 Excel 的宿主测试。

.EXAMPLE
.\scripts\verify-all.ps1 -SkipHost
跳过宿主测试，适合改完纯逻辑后快速回归。
#>
[CmdletBinding()]
param(
    [ValidateSet('Debug', 'Release')]
    [string]$Configuration = 'Release',

    [switch]$SkipHost,

    [switch]$IncludePanelUi
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
# 控制台默认按系统 ANSI 代码页输出，中文测试名会变成乱码，判定结果无从对照。
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$RepoRoot = Split-Path -Parent $PSScriptRoot
$LogRoot  = Join-Path $RepoRoot 'artifacts\verify-all'
$Solution = Join-Path $RepoRoot 'ChatSheet.sln'
$TestExe  = Join-Path $RepoRoot "tests\ChatSheet.ToolTests\bin\$Configuration\ChatSheet.ToolTests.exe"
$PaneExe  = Join-Path $RepoRoot "tests\ChatSheet.PaneHarness\bin\$Configuration\ChatSheet.PaneHarness.exe"
$WebTests = Join-Path $RepoRoot 'tests\web'

function Write-Step { param([string]$T) Write-Host ""; Write-Host "==> $T" -ForegroundColor Cyan }
function Write-Ok   { param([string]$T) Write-Host "    $T" -ForegroundColor Green }
function Write-Bad  { param([string]$T) Write-Host "    $T" -ForegroundColor Red }
function Write-Note { param([string]$T) Write-Host "    $T" -ForegroundColor Yellow }

# StrictMode 下没运行过任何原生命令时，读 $LASTEXITCODE 会直接抛「变量未设置」，
# 所以统一走这个取值函数，缺省按失败处理而不是让脚本崩在中途。
function Get-LastExitCode {
    if (Test-Path variable:global:LASTEXITCODE) { return [int]$global:LASTEXITCODE }
    return -1
}

# 日志名要能直接当文件名，又得保留中文步骤名，只替换文件系统不认的字符。
function Get-SafeName {
    param([string]$Name)
    return ($Name -replace '[\\/:*?"<>|]', '_')
}

$Results = New-Object System.Collections.Generic.List[object]

function Invoke-VerifyStep {
    param(
        [string]$Name,
        [string]$FilePath,
        [string[]]$Arguments = @(),
        [switch]$Skip,
        [string]$SkipReason = ''
    )

    if ($Skip) {
        Write-Step "$Name（跳过：$SkipReason）"
        $Results.Add([pscustomobject]@{ Name = $Name; State = '跳过'; Code = 0; Detail = $SkipReason; Log = '' })
        return
    }

    Write-Step $Name
    if (-not (Test-Path -LiteralPath $FilePath)) {
        Write-Bad "找不到可执行文件：$FilePath"
        $Results.Add([pscustomobject]@{ Name = $Name; State = '失败'; Code = -1; Detail = "缺少 $FilePath"; Log = '' })
        return
    }

    $log = Join-Path $LogRoot ((Get-SafeName -Name $Name) + '.log')
    # 外部程序把 stderr 写到控制台时，PowerShell 5.1 在 $ErrorActionPreference='Stop'
    # 下会把它升级成终止错误，测试进程自己的退出码就读不到了。
    # 因此调用期间临时降到 Continue，退出码仍然由 $LASTEXITCODE 如实反映。
    $previous = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        & $FilePath @Arguments *>&1 | Tee-Object -FilePath $log | Out-Host
        $code = Get-LastExitCode
    }
    finally {
        $ErrorActionPreference = $previous
    }

    if ($code -eq 0) {
        Write-Ok '通过'
        $Results.Add([pscustomobject]@{ Name = $Name; State = '通过'; Code = 0; Detail = ''; Log = $log })
    }
    else {
        Write-Bad "失败（退出码 $code），日志：$log"
        $Results.Add([pscustomobject]@{ Name = $Name; State = '失败'; Code = $code; Detail = "退出码 $code"; Log = $log })
    }
}

if (-not (Test-Path -LiteralPath $Solution)) {
    throw "不在仓库根目录运行：找不到 $Solution"
}

foreach ($tool in @('dotnet', 'node')) {
    if (-not (Get-Command $tool -ErrorAction SilentlyContinue)) {
        throw "缺少 $tool，无法验证。构建需要 .NET SDK，web 单测需要 Node.js。"
    }
}

if (Test-Path -LiteralPath $LogRoot) {
    # 上一次的日志混进来会让「看日志定位本次失败」变成翻旧账。
    Remove-Item -LiteralPath $LogRoot -Recurse -Force
}
New-Item -ItemType Directory -Path $LogRoot -Force | Out-Null

Write-Host '========================================' -ForegroundColor Magenta
Write-Host " Office-helper 全量验证（$Configuration）" -ForegroundColor Magenta
Write-Host '========================================' -ForegroundColor Magenta
Write-Note "仓库：$RepoRoot"
if ($SkipHost) { Write-Note '已指定 -SkipHost：跳过需要 Excel 的宿主测试' }

Invoke-VerifyStep -Name '构建 ChatSheet.sln' -FilePath (Get-Command dotnet).Source -Arguments @('build', $Solution, '--configuration', $Configuration, '--nologo')

# 构建日志单独再扫一遍警告数：0 警告是本仓库的验收线，但警告不该让验证红掉——
# 真正阻断发布的是编译错误，警告作为提示留给提交者判断。
$buildLog = Join-Path $LogRoot '构建 ChatSheet.sln.log'
if ((Test-Path -LiteralPath $buildLog) -and $Results[0].State -eq '通过') {
    $warningCount = @(Select-String -LiteralPath $buildLog -Pattern 'warning ' -SimpleMatch).Count
    if ($warningCount -gt 0) { Write-Note "构建警告 $warningCount 条（验收线是 0 警告）" }
}

Invoke-VerifyStep -Name 'Agent 工具回归 --agent-tool-tests' -FilePath $TestExe -Arguments @('--agent-tool-tests')
Invoke-VerifyStep -Name '可靠性回归 --reliability-tests' -FilePath $TestExe -Arguments @('--reliability-tests')
Invoke-VerifyStep -Name '项目审查 --project-audit-tests' -FilePath $TestExe -Arguments @('--project-audit-tests') -Skip:$SkipHost -SkipReason '需要真实 Excel'
Invoke-VerifyStep -Name 'Codex 配置 --codex-config-tests' -FilePath $TestExe -Arguments @('--codex-config-tests')

# web 单测每个文件独立进程：一个文件里 process.exit 或全局状态污染不会连坐其他文件。
# 数量按目录实际内容算，不写死 37——新增用例时脚本不需要跟着改。
$webFiles = @(Get-ChildItem -LiteralPath $WebTests -Filter '*.test.mjs' -File | Sort-Object Name)
if ($webFiles.Count -eq 0) {
    Write-Bad "没有找到 web 单测：$WebTests\*.test.mjs"
    $Results.Add([pscustomobject]@{ Name = 'web 单测'; State = '失败'; Code = -1; Detail = '未找到 *.test.mjs'; Log = '' })
}
else {
    Write-Step "web 单测：$($webFiles.Count) 个文件"
    $webFailed = New-Object System.Collections.Generic.List[string]
    foreach ($file in $webFiles) {
        $log = Join-Path $LogRoot ('web.' + [System.IO.Path]::GetFileNameWithoutExtension($file.Name) + '.log')
        $previous = $ErrorActionPreference
        $ErrorActionPreference = 'Continue'
        try {
            & (Get-Command node).Source $file.FullName *>&1 | Out-File -LiteralPath $log -Encoding utf8
            $code = Get-LastExitCode
        }
        finally {
            $ErrorActionPreference = $previous
        }

        if ($code -eq 0) {
            Write-Host "    PASS $($file.Name)" -ForegroundColor DarkGreen
            $Results.Add([pscustomobject]@{ Name = "web: $($file.Name)"; State = '通过'; Code = 0; Detail = ''; Log = $log })
        }
        else {
            Write-Host "    FAIL $($file.Name)（退出码 $code）" -ForegroundColor Red
            $webFailed.Add($file.Name)
            $Results.Add([pscustomobject]@{ Name = "web: $($file.Name)"; State = '失败'; Code = $code; Detail = "退出码 $code"; Log = $log })
        }
    }
    if ($webFailed.Count -gt 0) { Write-Bad "web 单测失败 $($webFailed.Count) 个：$($webFailed -join '、')" }
}

if ($IncludePanelUi) {
    Invoke-VerifyStep -Name '面板安全 --panel-security' -FilePath $PaneExe -Arguments @('--panel-security')
    Invoke-VerifyStep -Name 'WorkBuddy 面板 UI --workbuddy-ui' -FilePath $PaneExe -Arguments @('--workbuddy-ui', '--width', '360')
}

# ---- 汇总 ----
$passed = @($Results | Where-Object { $_.State -eq '通过' })
$failed = @($Results | Where-Object { $_.State -eq '失败' })
$skipped = @($Results | Where-Object { $_.State -eq '跳过' })

Write-Host ''
Write-Host '========================================' -ForegroundColor Magenta
Write-Host "汇总：通过 $($passed.Count)，失败 $($failed.Count)，跳过 $($skipped.Count)" -ForegroundColor $(if ($failed.Count -gt 0) { 'Red' } else { 'Green' })
Write-Host '========================================' -ForegroundColor Magenta

if ($skipped.Count -gt 0) {
    Write-Note '跳过项（去掉 -SkipHost 可一并验证）：'
    foreach ($s in $skipped) { Write-Host "    - $($s.Name)：$($s.Detail)" -ForegroundColor Yellow }
}

if ($failed.Count -gt 0) {
    Write-Bad '失败清单：'
    foreach ($f in $failed) {
        $where = if ($f.Log) { $f.Log } else { '（无日志）' }
        Write-Host "    - $($f.Name)：$($f.Detail)  →  $where" -ForegroundColor Red
    }
    Write-Host ''
    Write-Host "日志目录：$LogRoot" -ForegroundColor Yellow
    exit 1
}

Write-Host ''
Write-Host '全部通过。' -ForegroundColor Green
Write-Host "日志目录：$LogRoot" -ForegroundColor DarkGray
exit 0
