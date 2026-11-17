<#
.SYNOPSIS
Download and install an Office-helper Windows release.

.DESCRIPTION
Download the release ZIP and SHA-256 sidecar, verify, extract, and install.
-Version accepts a version with or without v. -KeepDownload saves both files.
Keep this bootstrap ASCII without a BOM so PowerShell 5.1 supports both -File
and the raw GitHub response piped to Invoke-Expression.
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
    if (-not $asset -or -not $hashAsset) { throw "Release $($release.tag_name) is missing the Windows ZIP or SHA-256 file." }

    $zipPath = Join-Path $workRoot $asset.name
    $hashPath = Join-Path $workRoot $hashAsset.name
    Invoke-WebRequest -Uri $asset.browser_download_url -Headers $headers -OutFile $zipPath -UseBasicParsing
    Invoke-WebRequest -Uri $hashAsset.browser_download_url -Headers $headers -OutFile $hashPath -UseBasicParsing

    $expected = ((Get-Content -LiteralPath $hashPath -Raw) -split '\s+')[0].ToLowerInvariant()
    $actual = (Get-FileHash -LiteralPath $zipPath -Algorithm SHA256).Hash.ToLowerInvariant()
    if ($actual -ne $expected) { throw 'SHA-256 verification failed: the download does not match the release.' }

    $extractRoot = Join-Path $workRoot 'extracted'
    Expand-Archive -LiteralPath $zipPath -DestinationPath $extractRoot -Force
    $package = Join-Path $extractRoot "ChatSheet-v$versionName-win"
    $installer = Join-Path $package 'scripts\install.ps1'
    if (-not (Test-Path -LiteralPath $installer)) { throw 'Invalid release package layout: scripts\install.ps1 is missing.' }
    Write-Host "Verified Office-helper $($release.tag_name); starting install or update." -ForegroundColor Green
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $installer -Action install
    if ($LASTEXITCODE -ne 0) { throw "Installer failed with exit code $LASTEXITCODE." }
    if ($KeepDownload) {
        Copy-Item -LiteralPath $zipPath -Destination (Join-Path (Get-Location) $asset.name) -Force
        Copy-Item -LiteralPath $hashPath -Destination (Join-Path (Get-Location) $hashAsset.name) -Force
    }
}
finally {
    $safeRoot = [IO.Path]::GetFullPath($workRoot)
    $tempPrefix = [IO.Path]::GetFullPath($env:TEMP).TrimEnd('\') + '\ChatSheet-install-'
    if (-not $safeRoot.StartsWith($tempPrefix, [StringComparison]::OrdinalIgnoreCase)) { throw 'Refusing to remove a path outside the installer temporary directory.' }
    Remove-Item -LiteralPath $safeRoot -Recurse -Force -ErrorAction SilentlyContinue
}
