<# 在线安装器离线回归：使用本地 ZIP，不下载、不注册 COM、不修改已安装的加载项。 #>
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$installer = [scriptblock]::Create([IO.File]::ReadAllText((Join-Path $repoRoot 'scripts\install-online.ps1')))
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('OfficeHelper-installer-tests-' + [Guid]::NewGuid().ToString('N'))
$packageName = 'ChatSheet-v0.10.3.21-win'
$assetName = "$packageName.zip"
$passed = 0
$failed = 0
$state = @{}

function Invoke-RestMethod {
    param($Uri, $Headers)
    $state.ReleaseUri = $Uri
    return $state.Release
}

function Invoke-WebRequest {
    param($Uri, $Headers, $OutFile, [switch]$UseBasicParsing)
    $state.BasicParsing = $state.BasicParsing -and $UseBasicParsing.IsPresent
    $state.WorkRoot = Split-Path -Parent $OutFile
    $source = if ($Uri -eq 'fixture:zip') { $state.Zip } else { $state.Hash }
    Copy-Item -LiteralPath $source -Destination $OutFile
}

function Check {
    param([string]$Label, [scriptblock]$Test)
    try {
        & $Test
        $script:passed++
        Write-Host "通过 $Label"
    } catch {
        $script:failed++
        Write-Host "失败 ${Label}：$($_.Exception.Message)"
    }
}

function Assert {
    param([bool]$Ok, [string]$Message)
    if (-not $Ok) { throw $Message }
}

function Reset-Fixture {
    $state.Clear()
    $state.Zip = Join-Path $testRoot $assetName
    $state.Hash = "$($state.Zip).sha256"
    $state.BasicParsing = $true
    $state.WorkRoot = $null
    $state.Release = [pscustomobject]@{
        tag_name = 'v0.10.3.21'
        assets = @(
            [pscustomobject]@{ name = $assetName; browser_download_url = 'fixture:zip' },
            [pscustomobject]@{ name = "$assetName.sha256"; browser_download_url = 'fixture:hash' }
        )
    }
    $hash = (Get-FileHash -LiteralPath $state.Zip -Algorithm SHA256).Hash.ToLowerInvariant()
    [IO.File]::WriteAllText($state.Hash, "$hash  $assetName`n")
    foreach ($path in @($marker, (Join-Path $downloadRoot $assetName), (Join-Path $downloadRoot "$assetName.sha256"))) {
        if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force }
    }
}

function Expect-Failure {
    param([string]$Message)
    $errorMessage = $null
    try { & $installer } catch { $errorMessage = $_.Exception.Message }
    Assert ($null -ne $errorMessage -and $errorMessage.Contains($Message)) "未返回预期错误：$errorMessage"
    Assert (-not (Test-Path -LiteralPath $marker)) '失败前执行了包内安装脚本。'
    if ($state.WorkRoot) { Assert (-not (Test-Path -LiteralPath $state.WorkRoot)) '失败后遗留临时下载目录。' }
}

try {
    $packageRoot = Join-Path $testRoot $packageName
    $scriptsRoot = Join-Path $packageRoot 'scripts'
    $downloadRoot = Join-Path $testRoot 'downloads'
    New-Item -ItemType Directory -Path $scriptsRoot, $downloadRoot -Force | Out-Null
    $marker = Join-Path $testRoot 'installed.txt'
    $escapedMarker = $marker.Replace("'", "''")
    [IO.File]::WriteAllText((Join-Path $scriptsRoot 'install.ps1'), "param([string]`$Action)`r`n[IO.File]::WriteAllText('$escapedMarker', `$Action)`r`nexit 0`r`n")
    Compress-Archive -LiteralPath $packageRoot -DestinationPath (Join-Path $testRoot $assetName)
    Push-Location $downloadRoot
    try {
        Check 'latest 校验、解压并调用安装入口，清理临时目录' {
            Reset-Fixture
            & $installer
            Assert ($state.ReleaseUri -eq 'https://api.github.com/repos/aEboli/Office-helper/releases/latest') '未查询最新发行版。'
            Assert ((Get-Content -LiteralPath $marker -Raw) -eq 'install') '未调用安装动作。'
            Assert (-not (Test-Path -LiteralPath $state.WorkRoot)) '成功后遗留临时目录。'
        }
        Check 'Windows PowerShell 下载不依赖 IE 引擎' {
            Reset-Fixture
            & $installer
            Assert $state.BasicParsing '下载未指定 UseBasicParsing。'
        }
        foreach ($version in @('0.10.3.21', 'v0.10.3.21')) {
            Check "指定版本 $version 使用正确标签" {
                Reset-Fixture
                & $installer -Version $version
                Assert ($state.ReleaseUri -eq 'https://api.github.com/repos/aEboli/Office-helper/releases/tags/v0.10.3.21') "标签不正确：$($state.ReleaseUri)"
            }
        }
        Check 'KeepDownload 保留 ZIP 和校验文件并清理临时目录' {
            Reset-Fixture
            & $installer -KeepDownload
            Assert (Test-Path -LiteralPath (Join-Path $downloadRoot $assetName)) '未保留 ZIP。'
            Assert (Test-Path -LiteralPath (Join-Path $downloadRoot "$assetName.sha256")) '未保留 SHA-256 文件。'
            Assert (-not (Test-Path -LiteralPath $state.WorkRoot)) '保留下载时遗留临时目录。'
        }
        Check '哈希不一致时拒绝安装' {
            Reset-Fixture
            [IO.File]::WriteAllText($state.Hash, ('0' * 64) + "  $assetName")
            Expect-Failure 'SHA-256'
        }
        Check '缺少校验资产时拒绝安装' {
            Reset-Fixture
            $state.Release.assets = @($state.Release.assets[0])
            Expect-Failure 'SHA-256'
        }
        Check '拒绝发行包内不匹配的根目录' {
            Reset-Fixture
            $state.Release.tag_name = 'v0.10.3.22'
            $state.Release.assets[0].name = 'ChatSheet-v0.10.3.22-win.zip'
            $state.Release.assets[1].name = 'ChatSheet-v0.10.3.22-win.zip.sha256'
            Expect-Failure '发行包结构不完整'
        }
        Check '包内安装失败时报告退出码并清理临时目录' {
            Reset-Fixture
            $fixtureInstaller = Join-Path $scriptsRoot 'install.ps1'
            [IO.File]::WriteAllText($fixtureInstaller, [IO.File]::ReadAllText($fixtureInstaller).Replace('exit 0', 'exit 23'))
            $state.Zip = Join-Path $testRoot 'failed-install.zip'
            Compress-Archive -LiteralPath $packageRoot -DestinationPath $state.Zip
            $hash = (Get-FileHash -LiteralPath $state.Zip -Algorithm SHA256).Hash
            [IO.File]::WriteAllText($state.Hash, "$hash  $assetName")
            $errorMessage = $null
            try { & $installer } catch { $errorMessage = $_.Exception.Message }
            Assert ($null -ne $errorMessage -and $errorMessage.Contains('23')) "未报告安装失败：$errorMessage"
            Assert (-not (Test-Path -LiteralPath $state.WorkRoot)) '安装失败后遗留临时目录。'
        }
    } finally { Pop-Location }
} finally {
    $safeRoot = [IO.Path]::GetFullPath($testRoot)
    $tempPrefix = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\OfficeHelper-installer-tests-'
    if (-not $safeRoot.StartsWith($tempPrefix, [StringComparison]::OrdinalIgnoreCase)) { throw '拒绝清理测试目录之外的路径。' }
    Remove-Item -LiteralPath $safeRoot -Recurse -Force -ErrorAction SilentlyContinue
}
Write-Host "=== 在线安装器：通过 $passed，失败 $failed ==="
if ($failed -gt 0) { exit 1 }
