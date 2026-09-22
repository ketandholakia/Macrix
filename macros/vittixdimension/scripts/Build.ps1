param(
    [string]$SourceRoot = (Join-Path $PSScriptRoot ".."),
    [string]$OutputRoot = (Join-Path $PSScriptRoot "..\build"),
    [string]$ReleaseRoot = (Join-Path $PSScriptRoot "..\release\VittixDimensionTools")
)

$src = Join-Path $SourceRoot "src"
$forms = Join-Path $SourceRoot "forms"
$docs = Join-Path $SourceRoot "docs"
$gms = Join-Path $SourceRoot "GMS"

New-Item -ItemType Directory -Force -Path $OutputRoot | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $OutputRoot "src") | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $OutputRoot "forms") | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $OutputRoot "docs") | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $OutputRoot "GMS") | Out-Null
New-Item -ItemType Directory -Force -Path $ReleaseRoot | Out-Null

Copy-Item -Force (Join-Path $src "*.bas") (Join-Path $OutputRoot "src")
Copy-Item -Force (Join-Path $forms "*.frm") (Join-Path $OutputRoot "forms")
Copy-Item -Force (Join-Path $forms "*.frx") (Join-Path $OutputRoot "forms") -ErrorAction SilentlyContinue
Copy-Item -Force (Join-Path $docs "*.md") (Join-Path $OutputRoot "docs")
Copy-Item -Force (Join-Path $gms "*.gms") (Join-Path $OutputRoot "GMS")
Copy-Item -Force (Join-Path $src "*.bas") $ReleaseRoot
Copy-Item -Force (Join-Path $forms "*.frm") $ReleaseRoot
Copy-Item -Force (Join-Path $forms "*.frx") $ReleaseRoot -ErrorAction SilentlyContinue
Copy-Item -Force (Join-Path $docs "*.md") $ReleaseRoot
Copy-Item -Force (Join-Path $gms "*.gms") $ReleaseRoot

$manifestPath = Join-Path $OutputRoot "GMS\VittixDimensionTools.manifest.txt"
$moduleNames = Get-ChildItem -LiteralPath (Join-Path $OutputRoot "src") -Filter *.bas | Sort-Object Name | Select-Object -ExpandProperty Name
$formNames = Get-ChildItem -LiteralPath (Join-Path $OutputRoot "forms") -Filter * | Sort-Object Name | Select-Object -ExpandProperty Name
@(
    "Vittix Dimension Tools package",
    "",
    "Modules:",
    (($moduleNames | ForEach-Object { " - $_" }) -join [Environment]::NewLine),
    "",
    "Forms:",
    (($formNames | ForEach-Object { " - $_" }) -join [Environment]::NewLine),
    "",
    "Entry points:",
    " - VittixDimensionTools_Main",
    " - VittixDimensionTools_Create"
) -join [Environment]::NewLine | Set-Content -LiteralPath $manifestPath

$releaseManifestPath = Join-Path $ReleaseRoot "VittixDimensionTools.manifest.txt"
$releaseFiles = Get-ChildItem -LiteralPath $ReleaseRoot -File | Sort-Object Name | Select-Object -ExpandProperty Name
@(
    "Vittix Dimension Tools release package",
    "",
    "Files:",
    (($releaseFiles | ForEach-Object { " - $_" }) -join [Environment]::NewLine),
    "",
    "Entry points:",
    " - VittixDimensionTools_Main",
    " - VittixDimensionTools_Create"
) -join [Environment]::NewLine | Set-Content -LiteralPath $releaseManifestPath

Write-Host "Build complete: $OutputRoot"
Write-Host "Release package: $ReleaseRoot"
