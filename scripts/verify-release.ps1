<#
.SYNOPSIS
从 GitHub 侧核对 Release：资产在不在、下载得到的字节与本地是否同一份。

.DESCRIPTION
不信上传脚本自己的输出——它只证明请求返回了 201。真正要确认的是
「别人点那个链接能拿到东西，且拿到的就是我打的那个包」。
因此下载回来算 SHA-256，与本地产物逐字节比对。
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$Tag,
    [string]$Repo = 'aEboli/Office-helper',
    [string]$LocalDir = 'artifacts\release'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$passed = 0
$failed = 0
function Assert {
    param([string]$Label, [bool]$Ok, [string]$Detail = '')
    if ($Ok) { $script:passed++; Write-Host "  通过  $Label" -ForegroundColor Green }
    else { $script:failed++; Write-Host "  失败  $Label$(if ($Detail) { "：$Detail" })" -ForegroundColor Red }
}

function Get-GitHubToken {
    # 不能用管道喂 git：PowerShell 把字符串交给原生命令时按控制台编码重写行尾，
    # 而 git 的凭据解析器按换行切字段，收到回车就报
    # 「refusing to work with credential missing protocol field」。
    # 因此写一个只含换行符的临时文件，用 cmd 的重定向喂进去。
    # publish-release.ps1 里是同一个写法，两支必须一致。
    $tmp = [IO.Path]::GetTempFileName()
    try {
        [IO.File]::WriteAllText($tmp, "protocol=https`nhost=github.com`n`n", (New-Object Text.UTF8Encoding($false)))
        $out = & cmd /c "git credential fill < `"$tmp`"" 2>$null
    }
    finally {
        Remove-Item -LiteralPath $tmp -ErrorAction SilentlyContinue
    }

    foreach ($line in $out) {
        if ($line -like 'password=*') { return $line.Substring('password='.Length) }
    }

    throw '取不到 github.com 的凭据。git push 能成的话它应该在 Windows 凭据管理器里。'
}

$headers = @{
    Authorization          = "Bearer $(Get-GitHubToken)"
    Accept                 = 'application/vnd.github+json'
    'X-GitHub-Api-Version' = '2022-11-28'
    'User-Agent'           = 'ChatSheet-verify-script'
}

Write-Host "核对 $Repo 的 Release $Tag" -ForegroundColor Cyan
Write-Host ''

$release = Invoke-RestMethod -Uri "https://api.github.com/repos/$Repo/releases/tags/$Tag" `
    -Headers $headers -Method Get

Assert 'Release 存在' ($null -ne $release)
Assert '不是草稿（草稿别人看不到）' (-not $release.draft) "draft=$($release.draft)"
Assert '标了正式版而非预发布' (-not $release.prerelease) "prerelease=$($release.prerelease)"
Assert '正文非空' ($release.body.Length -gt 500) "正文 $($release.body.Length) 字符"
# tag 到底指向哪个提交。
#
# 此前这里只判 target_commitish 字面是不是 'main'，而那说明不了任何事：GitHub 建 tag
# 时若不指定落点就用默认分支**当时**的 HEAD，本地发布提交没推上去时 tag 会指向上一版
# 的提交——而上一版的提交确确实实在 main 上，于是那条断言照样通过。v0.10.0 第一次发布
# 就是这么错的：资产是 0.10.0，tag 指着 v0.9.1 的提交。
#
# 真检查只有一条：把 tag 解引用成提交，与本地 HEAD 比。附注 tag 要多解一层——
# refs/tags 拿到的是 tag 对象而不是提交，漏了这一步会拿 tag 对象的哈希去比提交哈希，
# 永远不等。
$tagRef = Invoke-RestMethod -Uri "https://api.github.com/repos/$Repo/git/ref/tags/$Tag" `
    -Headers $headers -Method Get
$tagCommit = if ($tagRef.object.type -eq 'tag') {
    (Invoke-RestMethod -Uri $tagRef.object.url -Headers $headers -Method Get).object.sha
}
else {
    $tagRef.object.sha
}

$localHead = (& git rev-parse HEAD 2>$null)
if ($localHead) { $localHead = $localHead.Trim() }

Assert 'tag 解引用得到一个提交' ($tagCommit -and $tagCommit.Length -eq 40) "$tagCommit"
Assert 'tag 指向本地这次发布的提交' ($tagCommit -eq $localHead) (
    "tag → $tagCommit，本地 HEAD → $localHead。不一致说明发布时本地提交还没推到远端，" +
    '于是 tag 落在了远端当时的 HEAD 上——按 tag 取源码会拿到上一个版本')

$expected = @("ChatSheet-$Tag-win.zip", "ChatSheet-$Tag-win.zip.sha256")
foreach ($name in $expected) {
    $asset = $release.assets | Where-Object { $_.name -eq $name }
    Assert "资产 $name 在" ($null -ne $asset)
    if (-not $asset) { continue }

    # 名字不能被 GitHub 追加成 xxx-1.zip：README 里写死了文件名。
    Assert "$name 名字未被改写" ($asset.name -eq $name) $asset.name

    $local = Join-Path $LocalDir $name
    if (-not (Test-Path -LiteralPath $local)) {
        Assert "$name 本地产物在（用于比对）" $false $local
        continue
    }

    $localSize = (Get-Item -LiteralPath $local).Length
    Assert "$name 大小与本地一致（$($asset.size) 字节）" ($asset.size -eq $localSize) `
        "远端 $($asset.size) vs 本地 $localSize"

    # 真的下一遍：上传返回 201 不等于别人点链接能拿到东西。
    $tmp = Join-Path $env:TEMP "verify-$name"
    Invoke-WebRequest -Uri $asset.browser_download_url -OutFile $tmp -UseBasicParsing
    $remoteHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $tmp).Hash
    $localHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $local).Hash
    Assert "$name 下载回来与本地逐字节相同" ($remoteHash -eq $localHash) `
        "远端 $($remoteHash.Substring(0,16))… vs 本地 $($localHash.Substring(0,16))…"
    Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
}

# 校验文件里记的哈希，要与 ZIP 实际的哈希对得上——否则用户照它校验会失败。
$sumAsset = $release.assets | Where-Object { $_.name -like '*.sha256' }
$zipAsset = $release.assets | Where-Object { $_.name -like '*.zip' }
if ($sumAsset -and $zipAsset) {
    # GitHub 把 .sha256 当二进制发（application/octet-stream），于是 .Content 是
    # byte[] 而不是字符串。直接 -split 会把它强转成字符串，得到第一个字节的十进制值
    # （'9' 是 57），看起来像「哈希对不上」。必须显式按 UTF-8 解码。
    $raw = (Invoke-WebRequest -Uri $sumAsset.browser_download_url -UseBasicParsing).Content
    $sumText = if ($raw -is [byte[]]) { [System.Text.Encoding]::UTF8.GetString($raw) } else { [string]$raw }
    $recorded = ($sumText.Trim() -split '\s+')[0]
    $zipLocal = Join-Path $LocalDir $zipAsset.name
    $actual = (Get-FileHash -Algorithm SHA256 -LiteralPath $zipLocal).Hash
    Assert '校验文件里的哈希与 ZIP 实际哈希一致' ($recorded -ieq $actual) `
        "记录 $($recorded.Substring(0, [Math]::Min(16, $recorded.Length)))… vs 实际 $($actual.Substring(0,16))…"
}

Write-Host ''
Write-Host "=== $Tag：通过 $passed，失败 $failed ===" -ForegroundColor $(if ($failed -eq 0) { 'Green' } else { 'Red' })
if ($failed -gt 0) { exit 1 }
