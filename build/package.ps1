# package.ps1 — assemble a versioned source package + docs for the installer.
# Requires Windows PowerShell. Produces build\_out\{version}\VittixCore.gms-all\ ...
# Reads version from src\modVersion.bas (VER_CurrentVersion).
param(
    [string]$Src = (Join-Path $PSScriptRoot '..\src'),
    [string]$Out = (Join-Path $PSScriptRoot '_out')
)
$ErrorActionPreference = 'Stop'

# parse VER_CurrentVersion "x.y.z"
$verFile = Join-Path $Src 'modVersion.bas'
if (Test-Path $verFile) {
    $m = Select-String -Path $verFile -Pattern 'VER_CurrentVersion\s*=\s*"([\d.]+)"'
    $version = if ($m) { $m.Matches[0].Groups[1].Value } else { '0.0.0' }
} else {
    $version = '0.0.0'
}

$staging = Join-Path $Out 'Vittix'
if (Test-Path $staging) { Remove-Item $staging -Recurse -Force }
New-Item -ItemType Directory -Path $staging -Force | Out-Null

# assemble source tree
Copy-Item -Path $Src -Destination $staging -Recurse -Force
Copy-Item -Path (Join-Path $PSScriptRoot '\installer') -Destination $staging -Recurse -Force -ErrorAction SilentlyContinue

$zip = Join-Path $Out ("Vittix" + ($version -replace '\.','') + '.zip')
if (Test-Path $zip) { Remove-Item $zip -Force }
Compress-Archive -Path (Join-Path $staging '*') -DestinationPath $zip
Write-Host "Packaged $version -> $zip"