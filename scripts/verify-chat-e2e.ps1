<#
.SYNOPSIS
端到端验证对话链路：流式文本、工具调用、工具执行、消息推送。

.DESCRIPTION
用本地 mock 服务替代真实接口，因此不消耗任何额度、不需要真实密钥。
会临时改写设置以指向 mock，结束后自动还原用户原有设置。

验证重点是线程模型：Agent 循环在 await 之后位于线程池线程，
而 WebView2 与宿主 COM 都要求 UI 线程，此前正是这里出错导致「发消息没反应」。
#>
[CmdletBinding()]
param(
    [int]$MockPort = 58940,

    # 默认策略是写操作逐项审批，这是用户实际会遇到的路径，需单独验证。
    [ValidateSet('Automatic', 'PerWrite', 'PerTurn')]
    [string]$Approval = 'Automatic',

    # 跳过部署，直接验证已安装的版本。仅在确认产物已同步时使用。
    [switch]$SkipDeploy,

    # 上下文预算。下限受 Settings.Normalize 约束为 8000，
    # 配合 bulk 场景可堆到 90% 阈值以验证压缩路径。
    [int]$ContextBudget = 100000,

    # mock 场景：tool 走一次工具调用后收尾；
    # bulk 连续多轮读取以快速堆高上下文；
    # image 只回文本并报出收到的图片数，用于验证多模态链路；
    # flaky 前两次以 503 拒绝，验证失败重试会重来并最终成功；
    # reject 一律以 401 拒绝，验证配置类错误不被重试；
    # cut 第一轮模拟输出被长度上限截断，验证加载项自动续跑而不是当成结束；
    # cutloop 每轮都被截断，验证续跑有上限、不会无限空转；
    # notool 带 tools 就以 400 拒绝；toolrefusal 接受 tools 但声称没有调用能力；
    # deepseek-flash 验证自动模式首轮直接用文本指令协议；三者都验证照样能动手；
    # novision 带图片就以 400 拒绝，验证视觉回退（去图或经中转转写）。
    [ValidateSet('tool', 'bulk', 'image', 'flaky', 'reject', 'cut', 'cutloop', 'notool', 'toolrefusal', 'deepseek-flash', 'novision', 'grant')]
    [string]$Scenario = 'tool',

    # image 与 novision 场景下附加一张测试图片。
    [switch]$WithImage,

    # novision 场景下配置视觉中转模型，验证转写路径而非去图路径。
    [string]$VisionRelayModel = '',

    [switch]$KeepOpen
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$RepoRoot = Split-Path -Parent $PSScriptRoot
$SettingsPath = Join-Path $env:LOCALAPPDATA 'ChatSheet\settings.json'
$SecretPath = Join-Path $env:LOCALAPPDATA 'ChatSheet\secrets\custom-api-token.bin'
$LogDir = Join-Path $env:LOCALAPPDATA 'ChatSheet\logs'
$Workbook = Join-Path $RepoRoot 'work\p0-test.xlsx'
$BackupSuffix = '.e2e-backup'

function Write-Step { param([string]$T) Write-Host "==> $T" -ForegroundColor Cyan }
function Write-Ok { param([string]$T) Write-Host "    $T" -ForegroundColor Green }
function Write-Bad { param([string]$T) Write-Host "    $T" -ForegroundColor Red }
function Write-Note { param([string]$T) Write-Host "    $T" -ForegroundColor Yellow }

function Backup-UserConfig {
    foreach ($p in @($SettingsPath, $SecretPath)) {
        if (Test-Path -LiteralPath $p) {
            Copy-Item -LiteralPath $p -Destination ($p + $BackupSuffix) -Force
            Write-Note "已备份 $(Split-Path $p -Leaf)"
        }
    }
}

function Restore-UserConfig {
    foreach ($p in @($SettingsPath, $SecretPath)) {
        $backup = $p + $BackupSuffix
        if (Test-Path -LiteralPath $backup) {
            Copy-Item -LiteralPath $backup -Destination $p -Force
            Remove-Item -LiteralPath $backup -Force
            Write-Note "已还原 $(Split-Path $p -Leaf)"
        }
        elseif (Test-Path -LiteralPath $p) {
            # 原本不存在的文件，测试期间新建的要删掉。
            Remove-Item -LiteralPath $p -Force
            Write-Note "已移除测试产生的 $(Split-Path $p -Leaf)"
        }
    }
}

$mockJob = $null
$excelStarted = $false

try {
    Write-Step '备份用户配置'
    Backup-UserConfig

    Write-Step "启动 mock 服务（端口 $MockPort）"
    $mockScript = Join-Path $RepoRoot 'tests\mock-provider\server.mjs'
    # 脚本路径必须加引号：仓库路径含空格，否则 node 会把它按空格截断。
    $mockJob = Start-Process -FilePath 'node' -ArgumentList "`"$mockScript`" $MockPort $Scenario" `
        -PassThru -WindowStyle Hidden -RedirectStandardOutput (Join-Path $env:TEMP 'chatsheet-mock.log')
    Start-Sleep -Seconds 2

    $probe = Test-NetConnection -ComputerName 127.0.0.1 -Port $MockPort -InformationLevel Quiet -WarningAction SilentlyContinue
    if (-not $probe) { throw "mock 服务未监听端口 $MockPort。" }
    Write-Ok 'mock 服务就绪'

    Write-Step '写入指向 mock 的设置'
    # 直接写设置文件与密钥：绕过界面以便脚本化。
    $settings = [ordered]@{
        mode = 'CustomApi'
        cliSource = 'Auto'
        customProtocol = 'openai-chat-completions'
        customBaseUrl = "http://127.0.0.1:$MockPort/v1"
        model = if ($Scenario -eq 'deepseek-flash') { 'deepseek-v4-flash' } else { 'mock-model' }
        thinking = 'Off'
        approval = $Approval
        maxOutputTokens = 8192
        contextBudgetTokens = $ContextBudget
        maxSteps = 40
        autoIncludeSelection = $true
        # 一律用自动探测：这些场景要验的正是探测本身。
        toolProtocol = 'Auto'
        visionRelayModel = $VisionRelayModel
    }
    $json = $settings | ConvertTo-Json -Depth 5
    [System.IO.File]::WriteAllText($SettingsPath, $json, (New-Object System.Text.UTF8Encoding($true)))

    # mock 不校验密钥，但配置校验要求它存在，用 DPAPI 写一个占位值。
    Add-Type -AssemblyName System.Security
    $bytes = [System.Text.Encoding]::UTF8.GetBytes('mock-token')
    $entropy = [System.Text.Encoding]::UTF8.GetBytes('ChatSheet.SecretStore.v1')
    $cipher = [System.Security.Cryptography.ProtectedData]::Protect(
        $bytes, $entropy, [System.Security.Cryptography.DataProtectionScope]::CurrentUser)
    New-Item -ItemType Directory -Path (Split-Path $SecretPath) -Force | Out-Null
    [System.IO.File]::WriteAllBytes($SecretPath, $cipher)
    Write-Ok "已指向 http://127.0.0.1:$MockPort/v1，模型 $($settings.model)"

    Write-Step '部署当前构建'
    # 必须先部署：本脚本只负责启动宿主，不会自动同步产物。
    # 漏掉这步会验证到上一次安装的旧版本，得出与代码不符的结论。
    if (-not $SkipDeploy) {
        & (Join-Path $PSScriptRoot 'install.ps1') -Action install -SkipBuild | Out-Null
        Write-Ok '已部署最新产物'
    }
    else {
        Write-Note '按要求跳过部署，将验证已安装的版本'
    }

    Write-Step '启动 Excel 并打开面板'
    if (Test-Path -LiteralPath $LogDir) { Remove-Item -LiteralPath $LogDir -Recurse -Force }
    & (Join-Path $PSScriptRoot 'verify-panel.ps1') -Route chat -KeepOpen | Out-Null
    $excelStarted = $true

    Write-Step '通过面板发送消息'
    # 用 WebView2 执行脚本模拟真实点击：直接填入输入框并触发发送，
    # 走的是与用户操作完全相同的代码路径。
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.Text;

public static class XlApp
{
    delegate bool EnumProc(IntPtr hwnd, IntPtr lparam);
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc cb, IntPtr lparam);
    [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr parent, EnumProc cb, IntPtr lparam);
    [DllImport("user32.dll")] static extern int GetWindowThreadProcessId(IntPtr hwnd, out int pid);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassNameW(IntPtr h, StringBuilder t, int m);
    [DllImport("oleacc.dll")] static extern int AccessibleObjectFromWindow(IntPtr h, uint id, ref Guid iid,
        [MarshalAs(UnmanagedType.IUnknown)] out object obj);

    static string Cls(IntPtr h) { var sb = new StringBuilder(256); GetClassNameW(h, sb, sb.Capacity); return sb.ToString(); }

    public static object Get(int pid)
    {
        object result = null;
        EnumWindows((hwnd, l) => {
            int p; GetWindowThreadProcessId(hwnd, out p);
            if (p != pid || Cls(hwnd) != "XLMAIN") return true;
            EnumChildWindows(hwnd, (child, l2) => {
                if (Cls(child) != "EXCEL7") return true;
                var iid = new Guid("00020400-0000-0000-C000-000000000046");
                object w;
                if (AccessibleObjectFromWindow(child, 0xFFFFFFF0, ref iid, out w) == 0 && w != null)
                {
                    result = w.GetType().InvokeMember("Application",
                        System.Reflection.BindingFlags.GetProperty, null, w, null);
                    return false;
                }
                return true;
            }, IntPtr.Zero);
            return result == null;
        }, IntPtr.Zero);
        return result;
    }
}
'@

    $proc = Get-Process -Name EXCEL | Where-Object { $_.MainWindowHandle -ne 0 } | Select-Object -First 1
    $app = [XlApp]::Get($proc.Id)
    $automation = $app.COMAddIns.Item('ChatSheet.AddIn').Object
    $automation.ShowPane('chat')
    Start-Sleep -Seconds 2

    if ($WithImage) {
        Write-Step '附加测试图片'
        # 1×1 像素的合法 PNG，来自官方文档示例。
        $tinyPng = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGP4z8AAAAMBAQDJ/pLvAAAAAElFTkSuQmCC'
        $attached = $automation.AttachImageForTest($tinyPng, 'tiny.png')
        Write-Ok "附加结果：$attached"
        if ($attached -notmatch '已附加') {
            Write-Bad '图片未成功附加'
        }
    }

    $sent = $automation.SendChatForTest('把表头写成名称和数量')
    Write-Ok "已投递测试消息：$sent"

    $logFile = Join-Path $LogDir 'addin-EXCEL.log'

    if ($Approval -eq 'PerWrite' -or $Approval -eq 'PerTurn') {
        Write-Step '等待审批卡片并点击「允许」'
        # 逐项审批策略下，Agent 会挂起等待用户决定。
        # 这里用脚本点击真实按钮，走与手工操作相同的路径。
        #
        # PerTurn 与 grant 场景要一直点到对话收尾：这两个的验证点正是
        # 「弹了几次卡」，提前退出会让计数变成脚本决定的，而不是执行器决定的。
        $keepClicking = ($Approval -eq 'PerTurn') -or ($Scenario -eq 'grant')
        $approved = $false
        $clickCount = 0
        $deadline = (Get-Date).AddSeconds($(if ($keepClicking) { 90 } else { 30 }))
        while ((Get-Date) -lt $deadline) {
            Start-Sleep -Seconds 2
            $clicked = $automation.ClickApprovalForTest($true)
            if ($clicked -match '已点击') {
                Write-Ok $clicked
                $approved = $true
                $clickCount++
                if (-not $keepClicking) { break }
            }
            elseif ($keepClicking -and (Test-Path -LiteralPath $logFile)) {
                $soFar = Get-Content -LiteralPath $logFile -Encoding UTF8 -Raw
                if ($soFar -match '对话结束|对话失败|Agent 运行失败') { break }
            }
        }

        if (-not $approved) { Write-Bad '未出现审批卡片' }
        if ($keepClicking) { Write-Ok "共点了 $clickCount 次「允许」" }
    }

    Write-Step '等待对话完成（最多 60 秒）'
    $deadline = (Get-Date).AddSeconds(60)
    $done = $false
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Seconds 2
        if (-not (Test-Path -LiteralPath $logFile)) { continue }
        $content = Get-Content -LiteralPath $logFile -Encoding UTF8 -Raw
        # 失败收尾也算结束，否则验证失败路径的场景要白等满 60 秒。
        # 「停止续跑」同理：反复截断后主动收手，不会再留下对话结束的记录。
        if ($content -match '对话结束|对话失败|Agent 运行失败|停止续跑') { $done = $true; break }
    }

    if (-not $done) { Write-Note '未在时限内见到对话结束记录' }

    Write-Step '日志'
    if (Test-Path -LiteralPath $logFile) {
        Get-Content -LiteralPath $logFile -Encoding UTF8 | ForEach-Object { Write-Host "      $_" }
    }

    Write-Step 'mock 服务日志'
    $mockLog = Join-Path $env:TEMP 'chatsheet-mock.log'
    if (Test-Path -LiteralPath $mockLog) {
        Get-Content -LiteralPath $mockLog | ForEach-Object { Write-Host "      $_" }
    }

    if ($Scenario -eq 'grant') {
        Write-Step '授权分档判定'
        # 这一段验的是执行器的分流，不是面板的渲染：日志里每条审批请求
        # 都带工具名，数它们就能看出「同表同类问了几次」。
        $logText = if (Test-Path -LiteralPath $logFile) {
            Get-Content -LiteralPath $logFile -Encoding UTF8 -Raw
        } else { '' }

        $asked = [regex]::Matches($logText, '审批请求[：:]\s*(\S+)') |
            ForEach-Object { $_.Groups[1].Value }
        if ($asked.Count -eq 0) {
            # 退回读工具名：不同版本的日志措辞可能不同，此时至少报出原始行。
            $asked = [regex]::Matches($logText, '需要确认[^\r\n]*') | ForEach-Object { $_.Value }
        }

        Write-Ok ("问过的操作：" + ($(if ($asked.Count) { $asked -join '、' } else { '（日志里没找到审批记录）' })))

        $formatAsked = @($asked | Where-Object {
            $_ -match 'format_range|set_number_format|fit_range'
        }).Count
        $structureAsked = @($asked | Where-Object { $_ -match 'add_worksheet' }).Count

        if ($Approval -eq 'PerTurn') {
            if ($formatAsked -eq 1) {
                Write-Ok "格式类只问了 1 次（同表同类共用一笔授权）"
            } else {
                Write-Bad "格式类问了 $formatAsked 次，期望 1 次——同表同类应共用一笔授权"
            }
        }

        if ($structureAsked -ge 1) {
            Write-Ok '结构操作单独问过（格式授权没有把它捎带放行）'
        } else {
            Write-Bad '结构操作没有单独问——格式授权把 add_worksheet 放行了'
        }

        # 四步都要真的执行到，否则上面的计数可能只是因为提前失败。
        if ($logText -match '新表' -or $logText -match 'add_worksheet') {
            Write-Ok '四步走完，含结构操作'
        } else {
            Write-Note '日志里未见结构操作痕迹，计数可能不完整'
        }
    }

    if ($Scenario -eq 'image') {
        Write-Step '图片链路判定'
        $mockLogPath = Join-Path $env:TEMP 'chatsheet-mock.log'
        $mockText = if (Test-Path -LiteralPath $mockLogPath) {
            Get-Content -LiteralPath $mockLogPath -Raw
        } else { '' }

        $imageMatch = [regex]::Match($mockText, 'images=(\d+)(?: types=(\S+))?')
        if (-not $imageMatch.Success) {
            Write-Bad 'mock 未报告图片计数'
        }
        else {
            $count = [int]$imageMatch.Groups[1].Value
            $types = $imageMatch.Groups[2].Value
            Write-Ok "mock 收到 images=$count types=$types"

            if ($WithImage -and $count -ge 1) { Write-Ok '图片已送达服务端' }
            elseif ($WithImage) { Write-Bad '附加了图片但服务端未收到' }
            elseif ($count -eq 0) { Write-Ok '未附加图片时确实不发送图片块' }

            if ($WithImage -and $types -match 'image/png') { Write-Ok '媒体类型正确传递' }
        }

        # 加载项日志会记录本轮附带的图片数量。
        $addinLog = if (Test-Path -LiteralPath $logFile) {
            Get-Content -LiteralPath $logFile -Raw -Encoding UTF8
        } else { '' }

        if ($WithImage) {
            if ($addinLog -match '本轮附带 (\d+) 张图片') {
                Write-Ok "加载项记录：本轮附带 $($Matches[1]) 张图片"
            }
            else {
                Write-Bad '加载项日志未记录图片附带情况'
            }
        }

        if ($addinLog -match 'can only be accessed from the UI thread') {
            Write-Bad '存在 UI 线程访问错误'
        }
        else {
            Write-Ok '无 UI 线程访问错误'
        }

        return
    }

    # 重试场景自成一套判定：本轮只验证接口失败的处理，不写任何单元格，
    # 因此不走后面的撤销与写入核对（与 image 场景同样的结构）。
    if ($Scenario -in @('flaky', 'reject')) {
        Write-Step '失败重试判定'
        $log = if (Test-Path -LiteralPath $logFile) { Get-Content -LiteralPath $logFile -Encoding UTF8 -Raw } else { '' }
        if ($log -match 'can only be accessed from the UI thread') { Write-Bad '仍存在 UI 线程访问错误' }
        else { Write-Ok '无 UI 线程访问错误' }

        if ($log -match '开始对话') { Write-Ok '对话已发起' } else { Write-Bad '未见对话发起记录' }

        # 只看日志不够：必须数 mock 收到的请求次数，才能证明「真的重试了」
        # 与「没白重试」——否则一次都没重试也能靠日志文本蒙过去。
        $mockText = if (Test-Path -LiteralPath $mockLog) { Get-Content -LiteralPath $mockLog -Raw } else { '' }
        $attempts = @([regex]::Matches($mockText, '\[mock\] attempt (\d+) ->')).Count
        Write-Ok "mock 收到对话请求 $attempts 次"

        if ($Scenario -eq 'flaky') {
            # 前两次 503，第三次成功：共 3 次。
            if ($attempts -eq 3) { Write-Ok '失败两次后重试成功，请求次数符合预期' }
            else { Write-Bad "预期 3 次请求（失败 2 次 + 成功 1 次），实际 $attempts 次" }

            $retryLines = @([regex]::Matches($log, '接口调用失败，(\d+) 秒后重试（第 (\d+)/(\d+) 次）'))
            if ($retryLines.Count -ge 2) {
                $order = $retryLines | ForEach-Object { $_.Groups[2].Value }
                Write-Ok "重试日志 $($retryLines.Count) 条，次序 $($order -join '、')"
            }
            else { Write-Bad "预期至少 2 条重试记录，实际 $($retryLines.Count) 条" }

            # 带了 Retry-After: 1，首次重试应按它等 1 秒而非本地退避。
            if ($log -match '接口调用失败，1 秒后重试（第 1/5 次）') { Write-Ok '已采用服务端 Retry-After 的等待时长' }
            else { Write-Note '未见按 Retry-After 等待的记录' }

            if ($log -match '对话结束') { Write-Ok '重试后对话正常完成' } else { Write-Bad '重试后对话未完成' }
        }
        else {
            # 401 属配置错误，重试无意义，必须只请求一次就报错。
            if ($attempts -eq 1) { Write-Ok '配置类错误未被重试，只请求了一次' }
            else { Write-Bad "401 不应重试，但收到 $attempts 次请求" }

            if ($log -match '对话失败：HTTP_401.*接口返回 401') { Write-Ok '已记录并上报 401 原因' }
            else { Write-Bad '未见 401 错误上报' }
            if ($log -match '秒后重试') { Write-Bad '对 401 仍触发了重试' } else { Write-Ok '无重试记录' }
        }

        return
    }

    # 续跑上限判定：每轮都被截断时必须停下来，不能空转到步数上限。
    if ($Scenario -eq 'cutloop') {
        Write-Step '续跑上限判定'
        $log = if (Test-Path -LiteralPath $logFile) { Get-Content -LiteralPath $logFile -Encoding UTF8 -Raw } else { '' }
        $mockText = if (Test-Path -LiteralPath $mockLog) { Get-Content -LiteralPath $mockLog -Raw } else { '' }

        $continues = @([regex]::Matches($log, '自动继续（第 (\d+)/3 次）')).Count
        Write-Ok "自动续跑 $continues 次"
        if ($continues -eq 3) { Write-Ok '续跑次数正好用满上限' }
        else { Write-Bad "预期续跑 3 次，实际 $continues 次" }

        if ($log -match '连续 3 次输出被截断后仍无进展，停止续跑') { Write-Ok '已在上限处停止' }
        else { Write-Bad '未见停止续跑的记录' }

        # 关键：请求次数必须有界。4 次 = 首轮 + 3 次续跑。
        # 若这里等于步数上限（40），说明没停下来，空转烧了额度。
        $requests = @([regex]::Matches($mockText, 'cutloop: stalled turn')).Count
        Write-Ok "mock 收到对话请求 $requests 次"
        if ($requests -eq 4) { Write-Ok '请求次数有界，未空转到步数上限' }
        else { Write-Bad "预期 4 次请求（首轮 + 3 次续跑），实际 $requests 次" }

        if ($log -match '已达到单轮步数上限') { Write-Bad '空转到了步数上限，续跑上限没生效' }
        else { Write-Ok '未触及步数上限' }

        return
    }

    # 工具能力回退判定：无论是接口拒绝 tools，还是模型明确说没有调用能力，
    # 都应改用文本指令协议，并且解析出的调用要真的落到单元格里。
    if ($Scenario -in @('notool', 'toolrefusal', 'deepseek-flash')) {
        Write-Step '工具形态回退判定'
        $log = if (Test-Path -LiteralPath $logFile) { Get-Content -LiteralPath $logFile -Encoding UTF8 -Raw } else { '' }
        $mockText = if (Test-Path -LiteralPath $mockLog) { Get-Content -LiteralPath $mockLog -Raw } else { '' }

        if ($Scenario -eq 'notool') {
            if ($mockText -match 'tools not supported') { Write-Ok 'mock 已拒绝带 tools 的请求' }
            else { Write-Bad 'mock 未进入拒绝分支，本次没有验到回退' }
        }
        elseif ($Scenario -eq 'toolrefusal') {
            if ($mockText -match 'toolrefusal: model says it cannot call tools') { Write-Ok 'mock 已模拟模型声明没有工具调用能力' }
            else { Write-Bad 'mock 未模拟模型能力拒绝' }

            if ($mockText -match 'toolrefusal: retry omitted the refusal') { Write-Ok '重试上下文没有保留首次拒绝' }
            else { Write-Bad '首次拒绝仍留在重试上下文中' }

            if ($mockText -match 'toolrefusal: retry omitted native tools') { Write-Ok '文本协议重试未携带原生 tools' }
            else { Write-Bad '文本协议重试仍携带原生 tools' }
        }
        else {
            if ($mockText -match 'deepseek-flash: text protocol request \(tools=0\)') { Write-Ok 'DeepSeek Flash 首个请求直接使用文本协议' }
            else { Write-Bad 'DeepSeek Flash 首个请求没有确认使用文本协议' }

            if ($mockText -match 'deepseek-flash: unexpected native tools') { Write-Bad 'DeepSeek Flash 请求仍携带原生 tools' }
            else { Write-Ok 'DeepSeek Flash 请求未携带原生 tools' }
        }

        if ($Scenario -eq 'deepseek-flash') {
            if ($log -match '本轮工具形态：Text') { Write-Ok 'DeepSeek Flash 自动预选文本指令协议' }
            else { Write-Bad '未见 DeepSeek Flash 文本协议预选记录，整轮可能直接失败了' }
        }
        elseif ($log -match '工具形态降级为 Text') { Write-Ok '已改用文本指令协议' }
        else { Write-Bad '未见降级记录，整轮可能直接失败了' }

        # 降级后必须真的干成活。
        if ($log -match '工具 write_values 执行成功') { Write-Ok '文本协议下的调用已执行' }
        else { Write-Bad '文本协议下未见写入成功' }

        if ($log -match '对话结束') { Write-Ok '本轮正常收尾' } else { Write-Bad '本轮未正常收尾' }

        # 能力回退后原生模式只应出现一次，预先兼容时则一次都不应出现。
        if ($Scenario -eq 'notool') {
            $nativeAttempts = @([regex]::Matches($mockText, 'tools not supported')).Count
        }
        elseif ($Scenario -eq 'toolrefusal') {
            $nativeAttempts = @([regex]::Matches($mockText, 'toolrefusal: model says it cannot call tools')).Count
        }
        else {
            $nativeAttempts = @([regex]::Matches($mockText, 'tools=18')).Count
        }
        $expectedNativeAttempts = if ($Scenario -eq 'deepseek-flash') { 0 } else { 1 }
        if ($nativeAttempts -eq $expectedNativeAttempts) { Write-Ok "原生 tools 尝试次数符合预期：$nativeAttempts" }
        else { Write-Bad "预期原生 tools 尝试 $expectedNativeAttempts 次，实际 $nativeAttempts 次" }

        try {
            $sheet = $app.ActiveWorkbook.Worksheets.Item(1)
            $a1 = $sheet.Range('A1').Value2
            $b1 = $sheet.Range('B1').Value2
            if ($a1 -eq '名称' -and $b1 -eq '数量') { Write-Ok "单元格已写入：A1=$a1 B1=$b1" }
            else { Write-Bad "单元格内容不符：A1=$a1 B1=$b1" }
        }
        catch { Write-Bad "读取单元格失败：$_" }

        return
    }

    # 视觉回退判定：带图片被拒后应转写或去图，且必须告知，不能静默丢图。
    if ($Scenario -eq 'novision') {
        Write-Step '视觉回退判定'
        $log = if (Test-Path -LiteralPath $logFile) { Get-Content -LiteralPath $logFile -Encoding UTF8 -Raw } else { '' }
        $mockText = if (Test-Path -LiteralPath $mockLog) { Get-Content -LiteralPath $mockLog -Raw } else { '' }

        if (-not $WithImage) {
            Write-Note '未附加图片，本场景需要 -WithImage 才能验到回退'
            return
        }

        if ($mockText -match 'no vision') { Write-Ok 'mock 已拒绝带图片的请求' }
        else { Write-Bad 'mock 未进入拒绝分支，本次没有验到回退' }

        if ($log -match '不支持图片输入') { Write-Ok '已识别为模型无视觉能力' }
        else { Write-Bad '未见视觉能力判定记录' }

        if ($VisionRelayModel) {
            if ($log -match '视觉中转完成') { Write-Ok "已经过 $VisionRelayModel 转写" }
            else { Write-Bad '配置了中转模型但未见转写记录' }
        }

        # 关键：这一轮不能整轮失败，模型必须收到「有图但你看不到」的说明。
        if ($log -match '对话结束') { Write-Ok '去图后本轮仍正常收尾' }
        else { Write-Bad '视觉回退后本轮未收尾' }

        if ($mockText -match 'images=0') { Write-Ok '重发时确实不再带图片' }
        else { Write-Bad '重发仍带着图片，会反复撞同一个 400' }

        return
    }

    # 截断续跑自成一套判定：写入的是 A1:B1 而非 tool 场景的 A1:B2，
    # 因此不走后面按 tool 场景预期值核对的那段。
    if ($Scenario -eq 'cut') {
        Write-Step '截断自动续跑判定'
        $log = if (Test-Path -LiteralPath $logFile) { Get-Content -LiteralPath $logFile -Encoding UTF8 -Raw } else { '' }

        if ($log -match 'can only be accessed from the UI thread') { Write-Bad '仍存在 UI 线程访问错误' }
        else { Write-Ok '无 UI 线程访问错误' }

        # 核心断言：识别出无进展并自动续跑，而不是把半途当成本轮结束。
        if ($log -match '无进展（Truncated.*自动继续（第 1/3 次）') {
            Write-Ok '已识别输出截断并自动续跑'
        }
        else {
            Write-Bad '未见自动续跑记录，被截断的一步可能被当成正常结束'
        }

        # 续跑之后必须真的把活干完。
        if ($log -match '工具 write_values 执行成功') { Write-Ok '续跑后完成了写入' }
        else { Write-Bad '续跑后未见写入成功' }

        if ($log -match '对话结束') { Write-Ok '本轮正常收尾' } else { Write-Bad '本轮未正常收尾' }

        # mock 至少要收到 3 次请求：停顿轮、续跑轮、收尾轮。
        $mockText = if (Test-Path -LiteralPath $mockLog) { Get-Content -LiteralPath $mockLog -Raw } else { '' }
        if ($mockText -match 'cut: stalled turn') { Write-Ok 'mock 已模拟贴顶截断' }
        else { Write-Bad 'mock 未进入截断分支' }
        if ($mockText -match 'cut: continued, issuing tool call') { Write-Ok 'mock 收到了续跑请求' }
        else { Write-Bad 'mock 未收到续跑请求' }

        # 写入必须真的落到单元格里。
        try {
            $sheet = $app.ActiveWorkbook.Worksheets.Item(1)
            $a1 = $sheet.Range('A1').Value2
            $b1 = $sheet.Range('B1').Value2
            Write-Host "      A1 = $a1  B1 = $b1"
            if ("$a1" -eq '名称' -and "$b1" -eq '数量') { Write-Ok '续跑后的写入真实生效' }
            else { Write-Bad "单元格内容不符：A1=$a1 B1=$b1" }
        }
        catch {
            Write-Bad "读回单元格失败：$($_.Exception.Message)"
        }

        return
    }

    Write-Step '撤销与恢复'
    # 撤销必须真正回退单元格内容，而不只是界面变个状态。
    try {
        $sheet = $app.ActiveWorkbook.Worksheets.Item(1)
        $beforeUndo = $sheet.Range('A1').Value2

        $label = $automation.ClickUndoForTest(0)
        Start-Sleep -Seconds 2
        $afterUndo = $sheet.Range('A1').Value2
        Write-Ok "点击「$label」后 A1 = $(if ($null -eq $afterUndo) { '<空>' } else { $afterUndo })"

        if ($label -notmatch '撤销') {
            Write-Bad "首次点击的按钮文字应为「撤销」，实际为「$label」"
        }
        elseif ("$afterUndo" -ne "$beforeUndo") {
            Write-Ok "撤销已回退内容（$beforeUndo → $(if ($null -eq $afterUndo) { '<空>' } else { $afterUndo })）"
        }
        else {
            Write-Bad "撤销后 A1 仍为 $afterUndo，内容未回退"
        }

        $label2 = $automation.ClickUndoForTest(0)
        Start-Sleep -Seconds 2
        $afterRedo = $sheet.Range('A1').Value2
        Write-Ok "点击「$label2」后 A1 = $(if ($null -eq $afterRedo) { '<空>' } else { $afterRedo })"

        if ($label2 -notmatch '恢复') {
            Write-Bad "撤销后按钮文字应变为「恢复」，实际为「$label2」"
        }
        elseif ("$afterRedo" -eq "$beforeUndo") {
            Write-Ok '恢复已还原内容'
        }
        else {
            Write-Bad "恢复后 A1 为 $afterRedo，期望 $beforeUndo"
        }
    }
    catch {
        Write-Bad "撤销验证失败：$($_.Exception.Message)"
    }

    Write-Step '工作簿实际内容'
    # 工具报告成功不等于数据真的写进去了，必须读回单元格核对。
    try {
        $sheet = $app.ActiveWorkbook.Worksheets.Item(1)
        $cells = @{}
        foreach ($addr in 'A1', 'B1', 'A2', 'B2') {
            $cells[$addr] = $sheet.Range($addr).Value2
        }

        foreach ($addr in 'A1', 'B1', 'A2', 'B2') {
            Write-Host "      $addr = $($cells[$addr])"
        }

        $expected = @{ A1 = '名称'; B1 = '数量'; A2 = '铅笔'; B2 = 10 }
        $mismatch = @()
        foreach ($addr in $expected.Keys) {
            if ("$($cells[$addr])" -ne "$($expected[$addr])") {
                $mismatch += "$addr 期望 $($expected[$addr]) 实际 $($cells[$addr])"
            }
        }

        if ($mismatch.Count -eq 0) {
            Write-Ok '单元格内容与预期一致，写入真实生效'
        }
        else {
            Write-Bad ('单元格内容不符：' + ($mismatch -join '；'))
        }
    }
    catch {
        Write-Bad "读回单元格失败：$($_.Exception.Message)"
    }

    Write-Step '判定'
    $log = if (Test-Path -LiteralPath $logFile) { Get-Content -LiteralPath $logFile -Encoding UTF8 -Raw } else { '' }
    $threadError = $log -match 'can only be accessed from the UI thread'
    if ($threadError) { Write-Bad '仍存在 UI 线程访问错误' } else { Write-Ok '无 UI 线程访问错误' }
    if ($log -match '开始对话') { Write-Ok '对话已发起' } else { Write-Bad '未见对话发起记录' }
    if ($log -match '工具 \w+ 执行成功') { Write-Ok '工具已执行' } else { Write-Bad '未见工具执行记录' }
    if ($log -match '对话结束') { Write-Ok '对话已正常收尾' } else { Write-Bad '对话未正常收尾' }

    # 上下文圆环：每步都应推送一次，数值应随对话增长。
    # 日志形如「上下文圆环[路由进入]：444/100000 tokens = 0%」，来源标签可选。
    $ringLines = @([regex]::Matches($log, '上下文圆环(?:\[[^\]]*\])?：(\d+)/(\d+) tokens = (\d+)%'))
    if ($ringLines.Count -eq 0) {
        Write-Bad '上下文圆环未更新'
    }
    else {
        $values = $ringLines | ForEach-Object { [int]$_.Groups[1].Value }
        $grew = ($values | Select-Object -Last 1) -gt ($values | Select-Object -First 1)
        Write-Ok "上下文圆环更新 $($ringLines.Count) 次：$($values -join ' → ') tokens"
        if ($grew) { Write-Ok '圆环数值随对话增长' } else { Write-Note '圆环数值未增长（对话体量过小时属正常）' }
    }

    # 对话结束后的布局：工具卡片应默认折叠，消息应按 4/5 宽度分列。
    $chatLayout = [regex]::Match(
        $log,
        '对话布局：工具卡片 (\d+) 个（展开 (\d+)）\s*助手消息宽 (\S+)\s*用户消息宽 (\S+)\s*欢迎语 (\d+) 个')
    if (-not $chatLayout.Success) {
        Write-Note '未见对话布局上报'
    }
    else {
        $total = [int]$chatLayout.Groups[1].Value
        $opened = [int]$chatLayout.Groups[2].Value
        Write-Ok "工具卡片 $total 个，展开 $opened 个；助手宽 $($chatLayout.Groups[3].Value)，用户宽 $($chatLayout.Groups[4].Value)"

        if ($total -eq 0) { Write-Note '本轮未产生工具卡片' }
        elseif ($opened -eq 0) { Write-Ok '工具卡片默认折叠' }
        else { Write-Bad "有 $opened 个卡片默认就是展开的" }

        # 4/5 即 80%，允许边框与内边距带来的少量偏差。
        foreach ($pair in @(
            @{ Label = '助手消息'; Value = $chatLayout.Groups[3].Value },
            @{ Label = '用户消息'; Value = $chatLayout.Groups[4].Value })) {
            if ($pair.Value -match '^(\d+)%$') {
                $percent = [int]$Matches[1]
                if ($percent -le 82) { Write-Ok "$($pair.Label)宽度 $percent% 未超过 4/5" }
                else { Write-Bad "$($pair.Label)宽度 $percent% 超过了 4/5" }
            }
        }

        if ([int]$chatLayout.Groups[5].Value -ge 1) { Write-Ok '欢迎语已显示' }
        else { Write-Bad '欢迎语未显示' }
    }

    # 达到阈值时应记录压缩，并由界面提示。
    if ($Scenario -eq 'bulk') {
        if ($log -match '上下文压缩：') { Write-Ok '已触发上下文压缩' } else { Write-Bad '预算已调小但未触发压缩' }
        if ($log -match '已达阈值') { Write-Ok '圆环已标记达到阈值' } else { Write-Bad '圆环未标记达到阈值' }
    }

}
finally {
    if (-not $KeepOpen) {
        Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
    }

    if ($mockJob -and -not $mockJob.HasExited) {
        Stop-Process -Id $mockJob.Id -Force -ErrorAction SilentlyContinue
        Write-Note 'mock 服务已停止'
    }

    Write-Step '还原用户配置'
    Restore-UserConfig
}
