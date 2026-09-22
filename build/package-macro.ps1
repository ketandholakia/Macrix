<#
  build/package-macro.ps1 - assemble a release package for one registered macro.

  Reads the macro from macros\registry.json, copies its source/forms/docs into
  build\_out\macros\<id>\ and writes a MANIFEST.txt. Output lives under the
  already-gitignored build\_out tree, so generated artifacts never pollute the
  source folders (unlike the legacy per-macro scripts\Build.ps1).

  Usage:
    powershell -File build/package-macro.ps1 -Macro dimension-tools

  Exit codes: 0 ok | 6 unknown macro id
#>
param(
    [Parameter(Mandatory = $true)][string]$Macro,
    [string]$MacroRoot,
    [string]$OutRoot
)
$ErrorActionPreference = 'Stop'

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$regFile  = Join-Path $repoRoot 'macros\registry.json'
if (-not (Test-Path -LiteralPath $regFile)) { throw "macro registry not found: $regFile" }
$reg = Get-Content -Raw -LiteralPath $regFile -Encoding UTF8 | ConvertFrom-Json

$entry = $reg.macros | Where-Object { $_.id -eq $Macro }
if (-not $entry) {
    Write-Host "Unknown macro id '$Macro'. Run: powershell -File build\macros.ps1"
    exit 6
}

if (-not $MacroRoot) { $MacroRoot = $env:VITTIX_MACRO_ROOT }
if (-not $MacroRoot) { $MacroRoot = Join-Path $repoRoot $reg.macroRoot }
if (-not (Test-Path -LiteralPath $MacroRoot)) { throw "macro root not found: $MacroRoot" }

if (-not $OutRoot) { $OutRoot = Join-Path $repoRoot 'build\_out\macros' }
$out = Join-Path $OutRoot $entry.id

$mdir     = Join-Path $MacroRoot $entry.dir
$srcDir   = Join-Path $mdir 'src'
$formsDir = Join-Path $mdir 'forms'

if (Test-Path -LiteralPath $out) { Remove-Item -LiteralPath $out -Recurse -Force }
New-Item -ItemType Directory -Path $out -Force | Out-Null

$mods = @(); $frms = @(); $docs = @()
if (Test-Path -LiteralPath $srcDir)   { $mods = @(Get-ChildItem -LiteralPath $srcDir -File) }
if (Test-Path -LiteralPath $formsDir) { $frms = @(Get-ChildItem -LiteralPath $formsDir -File) }

$mods  | ForEach-Object { Copy-Item -LiteralPath $_.FullName -Destination $out -Force }
$frms  | ForEach-Object { Copy-Item -LiteralPath $_.FullName -Destination $out -Force }
$readme = Join-Path $mdir 'README.md'
if (Test-Path -LiteralPath $readme) { Copy-Item -LiteralPath $readme -Destination $out -Force; $docs += 'README.md' }

# ---- manifest (line list -- no nested arrays, so no "System.Object[]") ----
$lines = New-Object System.Collections.Generic.List[string]
$lines.Add("$($entry.display) - package")
$lines.Add("")
$lines.Add("Macro id    : $($entry.id)")
$lines.Add("GMS project : $($entry.project)")
$lines.Add("Stability   : $($entry.stability)")
$lines.Add("Imports forms: $($entry.importForms)")
$lines.Add("")
$lines.Add("Modules:")
if ($mods.Count -eq 0) { $lines.Add(" - (none)") }
foreach ($m in ($mods | Sort-Object Name)) { $lines.Add(" - $($m.Name)") }
$lines.Add("")
$lines.Add("Forms:")
if ($frms.Count -eq 0) { $lines.Add(" - (none)") }
foreach ($f in ($frms | Sort-Object Name)) { $lines.Add(" - $($f.Name)") }
$lines.Add("")
$lines.Add("Entry points:")
foreach ($e in $entry.entryPoints) { $lines.Add(" - $e") }
$lines.Add("")
$lines.Add("Built: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")

[System.IO.File]::WriteAllText((Join-Path $out 'MANIFEST.txt'), (($lines -join "`r`n") + "`r`n"), (New-Object System.Text.UTF8Encoding($false)))

Write-Host "Packaged '$Macro' -> $out"
Write-Host ("  modules={0} forms={1} docs={2}" -f $mods.Count, $frms.Count, $docs.Count)
exit 0
