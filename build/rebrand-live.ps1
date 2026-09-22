<#
  build/rebrand-live.ps1

  Finishes the Vittix -> Macrix rebrand for the *live* CorelDRAW projects, which
  cannot be done while CorelDRAW is running (a running instance rewrites its .gms
  files on exit and would undo the rename).

  Does three things:
    1. renames the .gms files in the CorelDRAW GMS folder,
    2. sets the frmDimension caption to "Macrix Dimension" in the live project,
    3. re-exports the form back into the repo.

  Afterwards run:  build\deploy.ps1 -Macro dimension-tools -Apply -StartCorel

  NOTE: the other three macros are renamed at the *project* level too. Any toolbar
  button, shortcut or script that refers to `VittixSelectSame.x`, `vittixBleed.x`
  or `vittixImposition.x` must be re-pointed at the Macrix names.
#>
$ErrorActionPreference = 'Stop'

$live = Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.ProcessName -match '(?i)corel' }
if ($live) {
    Write-Host 'CorelDRAW is running - close it first (its exit rewrites the .gms files).'
    exit 1
}

$gmsDir = $null
foreach ($root in @("$env:APPDATA\Corel", "$env:LOCALAPPDATA\Corel")) {
    if (-not (Test-Path $root)) { continue }
    $hit = Get-ChildItem -Path $root -Recurse -Directory -Filter 'GMS' -ErrorAction SilentlyContinue |
           Sort-Object FullName -Descending | Select-Object -First 1
    if ($hit) { $gmsDir = $hit.FullName; break }
}
if (-not $gmsDir) { Write-Host 'CorelDRAW GMS folder not found.'; exit 2 }
Write-Host ('GMS folder: ' + $gmsDir)

$renames = @(
    @('VittixDimensionTools.gms', 'MacrixTools.gms'),
    @('VittixSelectSame.gms',     'MacrixSelectSame.gms'),
    @('vittixBleed.gms',          'macrixBleed.gms'),
    @('vittixImposition.gms',     'macrixImposition.gms')
)
foreach ($r in $renames) {
    $from = Join-Path $gmsDir $r[0]
    $to = Join-Path $gmsDir $r[1]
    if (-not (Test-Path -LiteralPath $from)) { Write-Host ('  (absent) ' + $r[0]); continue }
    if (Test-Path -LiteralPath $to) { Write-Host ('  (target exists, skipped) ' + $r[1]); continue }
    Move-Item -LiteralPath $from -Destination $to -Force
    Write-Host ('  renamed ' + $r[0] + '  ->  ' + $r[1])
}

Write-Host ''
Write-Host 'Setting the form caption in the live project ...'
$app = $null
try {
    $app = New-Object -ComObject CorelDRAW.Application
    try { $app.Visible = $false } catch {}
    try { [void]$app.InitializeVBA() } catch {}
    Start-Sleep -Seconds 4
    $vbe = $app.VBE
    $target = $renames[0][1]
    $proj = $null
    foreach ($p in $vbe.VBProjects) {
        $fn = ''; try { $fn = $p.FileName } catch {}
        if ($fn -and (Split-Path $fn -Leaf) -ieq $target) { $proj = $p; break }
    }
    if (-not $proj) { Write-Host ('  project ' + $target + ' not found'); }
    else {
        Write-Host ('  project: ' + $proj.Name)
        foreach ($c in $proj.VBComponents) {
            if ($c.Type -eq 3 -and $c.Name -ieq 'frmDimension') {
                $c.Properties.Item('Caption').Value = 'Macrix Dimension'
                Write-Host ('  caption = ' + $c.Properties.Item('Caption').Value)
                $formsDir = Join-Path (Split-Path $PSScriptRoot '..') 'macros\macrixTools\forms'
                if (Test-Path -LiteralPath $formsDir) {
                    $c.Export((Join-Path $formsDir 'frmDimension.frm'))
                    Write-Host '  form re-exported'
                }
                $gp = $app.GMSManager.Projects.Item($proj.Name)
                Write-Host ('  macro list count: ' + $gp.Macros.Count)
            }
        }
    }
}
finally {
    if ($app) { try { $app.Quit() } catch {} }
}
Write-Host ''
Write-Host 'Next: powershell -File build\deploy.ps1 -Macro dimension-tools -Apply -StartCorel'
