param(
    [string]$SourceRoot = (Join-Path $PSScriptRoot ".."),
    [string]$TargetRoot = (Join-Path $env:APPDATA "Corel\CorelDRAW Graphics Suite 2021\Draw\GMS")
)

$releaseRoot = Join-Path $SourceRoot "release\MacrixTools"
$targetPackage = Join-Path $TargetRoot "MacrixTools"

function Assert-PathExists {
    param([string]$Path, [string]$Label)
    if (-not (Test-Path -LiteralPath $Path)) {
        throw "$Label path not found: $Path"
    }
}

Assert-PathExists -Path $releaseRoot -Label "Release"

New-Item -ItemType Directory -Force -Path $TargetRoot | Out-Null
New-Item -ItemType Directory -Force -Path $targetPackage | Out-Null

Copy-Item -Force (Join-Path $releaseRoot "*") $targetPackage -Exclude *.manifest.txt
Copy-Item -Force (Join-Path $releaseRoot "MacrixTools.manifest.txt") $TargetRoot -ErrorAction SilentlyContinue

Write-Host "Installed MacrixTools to $targetPackage"
