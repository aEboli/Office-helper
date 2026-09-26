<#
.SYNOPSIS
下载并安装最新的 ChatSheet Windows 发行包。

.DESCRIPTION
只下载 GitHub Release 的 ZIP 和 SHA-256 sidecar，校验通过后解压到临时目录，
再调用包内安装脚本。不会执行网络上直接管道下来的 PowerShell 代码。
#>
[CmdletBinding()]
param(
    [string]$Version = 'latest',
    [string]$Repository = 'aEboli/Office-helper',
    [switch]$KeepDownload
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$workRoot = Join-Path $env:TEMP ('ChatSheet-install-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $workRoot -Force | Out-Null
try {
    $headers = @{ 'User-Agent' = 'ChatSheet-online-installer'; Accept = 'application/vnd.github+json' }
    $release = if ($Version -eq 'latest') {
        Invoke-RestMethod -Uri "https://api.github.com/repos/$Repository/releases/latest" -Headers $headers
    } else {
        Invoke-RestMethod -Uri "https://api.github.com/repos/$Repository/releases/tags/v$Version" -Headers $headers
    }
    $versionName = $release.tag_name.TrimStart('v')
    $assetName = "ChatSheet-v$versionName-win.zip"
    $asset = $release.assets | Where-Object { $_.name -eq $assetName } | Select-Object -First 1
    $hashAsset = $release.assets | Where-Object { $_.name -eq "$assetName.sha256" } | Select-Object -First 1
    if (-not $asset -or -not $hashAsset) { throw "Release $($release.tag_name) 缺少 Windows ZIP 或 SHA-256 文件。" }

    $zipPath = Join-Path $workRoot $asset.name
    $hashPath = Join-Path $workRoot $hashAsset.name
    Invoke-WebRequest -Uri $asset.browser_download_url -Headers $headers -OutFile $zipPath
    Invoke-WebRequest -Uri $hashAsset.browser_download_url -Headers $headers -OutFile $hashPath

    $expected = ((Get-Content -LiteralPath $hashPath -Raw) -split '\s+')[0].ToLowerInvariant()
    $actual = (Get-FileHash -LiteralPath $zipPath -Algorithm SHA256).Hash.ToLowerInvariant()
    if ($actual -ne $expected) { throw "SHA-256 校验失败：下载内容与 Release 不一致。" }

    $extractRoot = Join-Path $workRoot 'extracted'
    Expand-Archive -LiteralPath $zipPath -DestinationPath $extractRoot -Force
    $package = Get-ChildItem -LiteralPath $extractRoot -Directory | Select-Object -First 1
    $installer = Join-Path $package.FullName 'scripts\install.ps1'
    if (-not (Test-Path -LiteralPath $installer)) { throw '发行包结构不完整，缺少安装脚本。' }
    Write-Host "已校验并解压 $($release.tag_name)，即将启动安装菜单。" -ForegroundColor Green
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $installer -Action install
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
    if ($KeepDownload) { Copy-Item -LiteralPath $zipPath -Destination (Join-Path (Get-Location) $asset.name) -Force }
}
finally {
    if (-not $KeepDownload) { Remove-Item -LiteralPath $workRoot -Recurse -Force -ErrorAction SilentlyContinue }
}
