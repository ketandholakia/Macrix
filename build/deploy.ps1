<#
  build/deploy.ps1 -- push repository VBA modules into a CorelDRAW VBA project (.gms).

  Grounded in the CorelDRAW 2021 automation model (verified against
  Programs64\TypeLibs\CorelDRAW.tlb): Application exposes
      VBE            -> VBIDE project/component model (import components)
      GMSManager     -> .Projects (.Load/.Unload), .RunMacro, .GMSPath
  So CorelDRAW can be driven from this script to import modules directly.

  TWO PATHS
    A. This script (-Apply): drives CorelDRAW over COM and imports components
       through Application.VBE, then attempts to save the target project.
       Requires CorelDRAW to permit VBA project-object-model access.
    B. Bootstrap (tools\dev-import\modDevImport.bas): stage the files with
       -Stage, import that ONE file into the project by hand once, then run
       DevImport_RunAll() inside CorelDRAW whenever you change source.

  SAFETY
    * Dry run by default: nothing is written unless you pass -Apply.
    * Only quits a CorelDRAW automation instance it launched itself (suppress with
      -KeepCorel); never touches a CorelDRAW you started normally.
    * Warns when the target project has unsaved changes.

  NOTE: a normally-launched CorelDRAW is not registered in the Running Object
  Table, so it cannot be attached to over COM. Use -StartCorel (launches an
  /Automation instance), or the bootstrap path (tools\dev-import\modDevImport.bas).

  EXIT CODES
    0 ok | 2 CorelDRAW not running | 3 VBE not accessible | 4 project not found
    5 validation failed

  EXAMPLES
    powershell -File build/deploy.ps1 -List
    powershell -File build/deploy.ps1
    powershell -File build/deploy.ps1 -Apply
    powershell -File build/deploy.ps1 -Stage
#>
[CmdletBinding()]
param(
    [string]$GmsDir,
    [string]$Project = 'vittix.gms',
    [switch]$List,
    [switch]$Apply,
    [switch]$Stage,
    [switch]$StartCorel,
    [switch]$KeepCorel,
    [switch]$SkipValidate
)

$ErrorActionPreference = 'Stop'
$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$srcDir   = Join-Path $repoRoot 'src'

function Get-GmsDir {
    foreach ($root in @("$env:APPDATA\Corel", "$env:LOCALAPPDATA\Corel")) {
        if (-not (Test-Path $root)) { continue }
        $hit = Get-ChildItem -Path $root -Recurse -Directory -Filter 'GMS' -ErrorAction SilentlyContinue |
               Sort-Object FullName -Descending | Select-Object -First 1
        if ($hit) { return $hit.FullName }
    }
    return $null
}

function Get-VBE {
    param($App)
    try { return $App.VBE } catch { return $null }
}

# ---------------------------------------------------------------- resolution
if (-not $GmsDir) { $GmsDir = Get-GmsDir }
if ($GmsDir) { Write-Host "GMS folder : $GmsDir" } else { Write-Host "GMS folder : <not found>" }

$modules = Get-ChildItem -LiteralPath $srcDir -File -ErrorAction SilentlyContinue |
           Where-Object { $_.Extension -in '.bas', '.cls', '.frm' } | Sort-Object Name
Write-Host ("Modules    : {0} in {1}" -f $modules.Count, $srcDir)

# ---------------------------------------------------------------- validation gate
if (-not $SkipValidate -and -not $List) {
    Write-Host "`n[1/4] Static validation ..."
    & (Join-Path $PSScriptRoot 'validate.ps1')
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Validation failed; aborting. (Use -SkipValidate to override.)"
        exit 5
    }
}

# ---------------------------------------------------------------- staging (path B)
if ($Stage) {
    $staging = Join-Path $env:APPDATA 'Vittix\staging'
    New-Item -ItemType Directory -Path $staging -Force | Out-Null
    $modules | ForEach-Object { Copy-Item -LiteralPath $_.FullName -Destination $staging -Force }
    Write-Host "`nStaged $($modules.Count) module(s) -> $staging"
    Write-Host "In CorelDRAW run:  modDevImport.DevImport_RunAll"
}

# ---------------------------------------------------------------- connect
Write-Host "`n[2/4] Connecting to CorelDRAW ..."
$app = $null
$owned = $false
# A normally-launched CorelDRAW is NOT in the Running Object Table, so it cannot
# be attached to; only an /Automation instance is. Try to attach, else optionally
# start our own automation instance (which we then own and may quit).
try { $app = [Runtime.InteropServices.Marshal]::GetActiveObject('CorelDRAW.Application') } catch { }
if (-not $app -and $StartCorel) {
    Write-Host "  launching a new CorelDRAW automation instance ..."
    $app = New-Object -ComObject CorelDRAW.Application
    $owned = $true
    try { $app.Visible = $false } catch { }
}
if (-not $app) {
    Write-Host "CorelDRAW is not reachable over COM."
    Write-Host "  A normally-launched CorelDRAW is not in the Running Object Table and cannot"
    Write-Host "  be attached to. Pass -StartCorel to launch an /Automation instance, or use the"
    Write-Host "  bootstrap path: build\deploy.ps1 -Stage  +  modDevImport.DevImport_RunAll."
    exit 2
}

$vbe = Get-VBE -App $app
if (-not $vbe) {
    Write-Host "CorelDRAW is running but Application.VBE is not accessible."
    Write-Host "Enable VBA project object-model access (Tools > Options > VBA in CorelDRAW),"
    Write-Host "then use the bootstrap path: build\deploy.ps1 -Stage  +  modDevImport.DevImport_RunAll."
    exit 3
}

# ---------------------------------------------------------------- -List
if ($List) {
    Write-Host "`n== Loaded GMS projects (GMSManager.Projects) =="
    try {
        $n = $app.GMSManager.Projects.Count
        for ($i = 1; $i -le $n; $i++) {
            $gp = $app.GMSManager.Projects.Item($i)
            Write-Host ("  {0,-24} {1}" -f $gp.Name, $gp.FullFileName)
            try {
                $mc = $gp.Macros.Count
                for ($j = 1; $j -le $mc; $j++) { Write-Host ("      - " + $gp.Macros.Item($j).Name) }
            } catch { }
        }
    } catch { Write-Host "  <GMSManager unavailable: $($_.Exception.Message)>" }

    Write-Host "`n== VBA projects (Application.VBE.VBProjects) =="
    foreach ($p in $vbe.VBProjects) { Write-Host ("  {0,-24} {1}" -f $p.Name, $p.FileName) }
    exit 0
}

# ---------------------------------------------------------------- find target project
Write-Host "`n[3/4] Locating target project '$Project' ..."
$proj = $null
foreach ($p in $vbe.VBProjects) {
    $leaf = ''
    try { if ($p.FileName) { $leaf = Split-Path $p.FileName -Leaf } } catch { }
    if ($leaf -ieq $Project) { $proj = $p; break }
}
if (-not $proj) {
    Write-Host "Target project '$Project' is not loaded in this CorelDRAW instance."
    Write-Host "Load it first (open the .gms / its VBA project), or run -List to see what is loaded."
    exit 4
}

# ---------------------------------------------------------------- plan
$plan = @()
foreach ($m in $modules) {
    $exists = $false
    foreach ($c in $proj.VBComponents) { if ($c.Name -ieq $m.BaseName) { $exists = $true; break } }
    $plan += [pscustomobject]@{
        Module = $m.BaseName
        Action = if ($exists) { 'Replace' } else { 'Add' }
        File   = $m.Name
    }
}
Write-Host "`nPlan for $($proj.Name) ($($plan.Count) module(s)):"
$plan | Format-Table -AutoSize | Out-String | Write-Host

$dirty = $false
try { $dirty = -not $proj.Saved } catch { }
if ($dirty) { Write-Host "WARNING: target project has unsaved edits; importing now may interact with them." }

# ---------------------------------------------------------------- apply
Write-Host "[4/4] " -NoNewline
if (-not $Apply) {
    Write-Host "DRY RUN -- nothing written. Re-run with -Apply to import."
    exit 0
}

Write-Host "applying ..."
$ok = 0; $fail = 0
foreach ($m in $modules) {
    try {
        $comp = $null
        foreach ($c in $proj.VBComponents) { if ($c.Name -ieq $m.BaseName) { $comp = $c; break } }
        if ($comp) { $proj.VBComponents.Remove($comp) }
        [void]$proj.VBComponents.Import($m.FullName)
        $ok++
        Write-Host ("  OK   " + $m.BaseName)
    } catch {
        $fail++
        Write-Host ("  FAIL " + $m.BaseName + " :: " + $_.Exception.Message)
    }
}

# ---------------------------------------------------------------- save
$saved = 'not attempted'
try {
    if ($proj.FileName) { $proj.SaveAs($proj.FileName); $saved = 'saved' }
    else { $saved = 'project has no FileName; save manually in the VBE' }
} catch { $saved = "save failed: $($_.Exception.Message)" }

Write-Host "`nImported $ok, failed $fail. Project save: $saved"
Write-Host "Note: a loaded .gms may also be persisted by CorelDRAW when the project is unloaded/closed - verify in the VBE."

if ($owned -and -not $KeepCorel) {
    try { $app.Quit() } catch { }
}

if ($fail -gt 0) { exit 1 }
exit 0
