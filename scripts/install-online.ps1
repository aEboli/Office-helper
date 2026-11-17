<#
.SYNOPSIS
下载并安装最新的 Office-helper Windows 发行包。

.DESCRIPTION
只下载 GitHub Release 的 ZIP 和 SHA-256 sidecar，校验通过后解压到临时目录，
再调用包内安装脚本。-Version 接受带或不带 v 的版本号；
-KeepDownload 将 ZIP 和校验文件保留在当前目录。
#>
[CmdletBinding()]
param(
    [string]$Version = 'latest',
    [string]$Repository = 'aEboli/Office-helper',
    [switch]$KeepDownload
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$Version = $Version.Trim()

$workRoot = Join-Path $env:TEMP ('ChatSheet-install-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $workRoot -Force | Out-Null
try {
    $headers = @{ 'User-Agent' = 'ChatSheet-online-installer'; Accept = 'application/vnd.github+json' }
    $release = if ($Version -eq 'latest') {
        Invoke-RestMethod -Uri "https://api.github.com/repos/$Repository/releases/latest" -Headers $headers
    } else {
        $tagVersion = $Version.TrimStart('v', 'V')
        Invoke-RestMethod -Uri "https://api.github.com/repos/$Repository/releases/tags/v$tagVersion" -Headers $headers
    }
    $versionName = $release.tag_name.TrimStart('v')
    $assetName = "ChatSheet-v$versionName-win.zip"
    $asset = $release.assets | Where-Object { $_.name -eq $assetName } | Select-Object -First 1
    $hashAsset = $release.assets | Where-Object { $_.name -eq "$assetName.sha256" } | Select-Object -First 1
    if (-not $asset -or -not $hashAsset) { throw "Release $($release.tag_name) 缺少 Windows ZIP 或 SHA-256 文件。" }

    $zipPath = Join-Path $workRoot $asset.name
    $hashPath = Join-Path $workRoot $hashAsset.name
    Invoke-WebRequest -Uri $asset.browser_download_url -Headers $headers -OutFile $zipPath -UseBasicParsing
    Invoke-WebRequest -Uri $hashAsset.browser_download_url -Headers $headers -OutFile $hashPath -UseBasicParsing

    $expected = ((Get-Content -LiteralPath $hashPath -Raw) -split '\s+')[0].ToLowerInvariant()
    $actual = (Get-FileHash -LiteralPath $zipPath -Algorithm SHA256).Hash.ToLowerInvariant()
    if ($actual -ne $expected) { throw "SHA-256 校验失败：下载内容与 Release 不一致。" }

    $extractRoot = Join-Path $workRoot 'extracted'
    Expand-Archive -LiteralPath $zipPath -DestinationPath $extractRoot -Force
    $package = Join-Path $extractRoot "ChatSheet-v$versionName-win"
    $installer = Join-Path $package 'scripts\install.ps1'
    if (-not (Test-Path -LiteralPath $installer)) { throw '发行包结构不完整，缺少安装脚本。' }
    Write-Host "已校验并解压 Office-helper $($release.tag_name)，即将安装或更新。" -ForegroundColor Green
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $installer -Action install
    if ($LASTEXITCODE -ne 0) { throw "安装程序失败，退出码 $LASTEXITCODE。" }
    if ($KeepDownload) {
        Copy-Item -LiteralPath $zipPath -Destination (Join-Path (Get-Location) $asset.name) -Force
        Copy-Item -LiteralPath $hashPath -Destination (Join-Path (Get-Location) $hashAsset.name) -Force
    }
}
finally {
    $safeRoot = [IO.Path]::GetFullPath($workRoot)
    $tempPrefix = [IO.Path]::GetFullPath($env:TEMP).TrimEnd('\') + '\ChatSheet-install-'
    if (-not $safeRoot.StartsWith($tempPrefix, [StringComparison]::OrdinalIgnoreCase)) { throw '拒绝清理安装临时目录之外的路径。' }
    Remove-Item -LiteralPath $safeRoot -Recurse -Force -ErrorAction SilentlyContinue
}
