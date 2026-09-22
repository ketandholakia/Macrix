<#
  build/macros.ps1 - inventory and verify the macros registered in this framework.

  Reads macros\registry.json (the declarative list of the CorelDRAW macro projects
  this framework manages). Sources live under a macro root, which defaults to the
  sibling checkout ..\macrixcdrMacro and can be overridden with -MacroRoot or the
  MACRIX_MACRO_ROOT environment variable.

  Usage:
    powershell -File build/macros.ps1              # table of macros + modules/forms
    powershell -File build/macros.ps1 -Json        # machine-readable dump
    powershell -File build/macros.ps1 -Verify      # exit 1 if any expected file is missing

  Exit codes: 0 ok | 1 verification problems | 2 macro root not found
#>
param(
    [string]$MacroRoot,
    [switch]$Verify,
    [switch]$Json
)
$ErrorActionPreference = 'Stop'

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$regPath  = Join-Path $repoRoot 'macros\registry.json'
if (-not (Test-Path -LiteralPath $regPath)) { throw "macro registry not found: $regPath" }
$reg = Get-Content -Raw -LiteralPath $regPath -Encoding UTF8 | ConvertFrom-Json

if (-not $MacroRoot) { $MacroRoot = $env:MACRIX_MACRO_ROOT }
if (-not $MacroRoot) { $MacroRoot = Join-Path $repoRoot $reg.macroRoot }
if (-not (Test-Path -LiteralPath $MacroRoot)) {
    Write-Host "Macro root not found: $MacroRoot"
    Write-Host "Set it with -MacroRoot <path> or the MACRIX_MACRO_ROOT environment variable."
    exit 2
}
$MacroRoot = (Resolve-Path -LiteralPath $MacroRoot).Path

$problems = @()
$rows = @()

foreach ($m in $reg.macros) {
    $dir    = Join-Path $MacroRoot $m.dir
    $srcDir = Join-Path $dir 'src'
    $frmDir = Join-Path $dir 'forms'

    if (-not (Test-Path -LiteralPath $dir)) {
        $problems += "$($m.id): macro folder not found ($($m.dir))"
        $rows += [pscustomobject]@{ Id = $m.id; Project = $m.project; Modules = 0; Forms = 0; EntryPoints = 0 }
        continue
    }

    $src = @(); if (Test-Path $srcDir) { $src = @(Get-ChildItem -LiteralPath $srcDir -File -ErrorAction SilentlyContinue) }
    $frm = @(); if (Test-Path $frmDir) { $frm = @(Get-ChildItem -LiteralPath $frmDir -File -ErrorAction SilentlyContinue) }

    # Every declared entry point must resolve to a procedure inside a src module.
    foreach ($ep in $m.entryPoints) {
        $parts = $ep -split '\.'
        $mod   = $parts[0]
        $proc  = if ($parts.Count -gt 1) { $parts[1] } else { $null }
        $hit = $src | Where-Object { $_.BaseName -ieq $mod }
        if (-not $hit) {
            $problems += "$($m.id): entry point '$ep' -> module '$mod' not found in $($m.dir)\src"
            continue
        }
        if ($proc) {
            $text = [System.IO.File]::ReadAllText($hit.FullName, [System.Text.Encoding]::UTF8)
            if ($text -notmatch "(?im)^\s*(Public\s+|Private\s+|Friend\s+)?(Sub|Function)\s+$([regex]::Escape($proc))\s*\(") {
                $problems += "$($m.id): entry point '$ep' -> procedure '$proc' not found in $($hit.Name)"
            }
        }
    }

    $rows += [pscustomobject]@{
        Id          = $m.id
        Project     = $m.project
        Modules     = @($src | Where-Object { $_.Extension -in '.bas', '.cls' }).Count
        Forms       = @($frm | Where-Object { $_.Extension -eq '.frm' }).Count
        EntryPoints = @($m.entryPoints).Count
    }
}

if ($Json) {
    [pscustomobject]@{ macroRoot = $MacroRoot; macros = $rows; problems = $problems } | ConvertTo-Json -Depth 5
}
else {
    Write-Host "Macro root: $MacroRoot"
    Write-Host ""
    $rows | Format-Table -AutoSize | Out-String | Write-Host
    if ($problems.Count -gt 0) {
        Write-Host "Problems:"
        $problems | ForEach-Object { Write-Host ("  ! " + $_) }
    }
    else {
        Write-Host "All registered macros resolved."
    }
}

if ($problems.Count -gt 0) { exit 1 }
exit 0
