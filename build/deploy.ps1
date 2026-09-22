<#
  build/deploy.ps1 - push/pull VBA modules between source and a CorelDRAW VBA project.

  Two targets:
    * this framework's own single project  (default: src/ -> <Project>)
    * a registered macro project           (-Macro <id>, from macros/registry.json)

  Grounded in the CorelDRAW 2021 automation model (verified against
  Programs64\TypeLibs\CorelDRAW.tlb): Application exposes
      VBE         -> VBIDE project/component model (import/export components)
      GMSManager  -> .Projects (.Load/.Unload), .RunMacro, .GMSPath

  DIRECTION
    (default)      plan a PUSH: source files -> project components
    -Apply         actually push
    -Pull          plan a PULL: project components -> source files
    -Pull -Apply   actually pull (overwrites source; git-tracked, so recoverable)
    -List          enumerate loaded GMS projects and their macros

  SAFETY
    * Dry run by default: nothing is written unless -Apply is passed.
    * Only quits a CorelDRAW automation instance it launched itself (-KeepCorel
      keeps even that); never touches a CorelDRAW you started normally.
    * Warns when the target project has unsaved edits.

  NOTE: a normally-launched CorelDRAW is not in the Running Object Table and
  cannot be attached to over COM. Use -StartCorel, or the bootstrap path
  (tools\dev-import\modDevImport.bas).

  EXIT CODES
    0 ok | 1 partial failures | 2 CorelDRAW unreachable | 3 VBE inaccessible
    4 project not found | 5 validation failed | 6 unknown macro id

  EXAMPLES
    powershell -File build/deploy.ps1 -List
    powershell -File build/deploy.ps1 -Macro dimension-tools
    powershell -File build/deploy.ps1 -Macro dimension-tools -Apply
    powershell -File build/deploy.ps1 -Macro dimension-tools -Pull -Apply
#>
[CmdletBinding()]
param(
    [string]$Macro,
    [string]$GmsDir,
    [string]$Project,
    [switch]$List,
    [switch]$Apply,
    [switch]$Pull,
    [switch]$Stage,
    [switch]$StartCorel,
    [switch]$KeepCorel,
    [switch]$SkipValidate
)

$ErrorActionPreference = 'Stop'
$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path

function Get-GmsDir {
    foreach ($root in @("$env:APPDATA\Corel", "$env:LOCALAPPDATA\Corel")) {
        if (-not (Test-Path $root)) { continue }
        $hit = Get-ChildItem -Path $root -Recurse -Directory -Filter 'GMS' -ErrorAction SilentlyContinue |
               Sort-Object FullName -Descending | Select-Object -First 1
        if ($hit) { return $hit.FullName }
    }
    return $null
}

function Get-VBE { param($App) try { return $App.VBE } catch { return $null } }

# ---------------------------------------------------------------- resolve target
$isMacro    = [bool]$Macro
$srcDir     = Join-Path $repoRoot 'src'      # framework project
$formsDir   = $null
$macroEntry = $null
$importForms = $false

if ($isMacro) {
    $regFile = Join-Path $repoRoot 'macros\registry.json'
    if (-not (Test-Path -LiteralPath $regFile)) { throw "macro registry not found: $regFile" }
    $mreg = Get-Content -Raw -LiteralPath $regFile -Encoding UTF8 | ConvertFrom-Json
    $macroEntry = $mreg.macros | Where-Object { $_.id -eq $Macro }
    if (-not $macroEntry) {
        Write-Host "Unknown macro id '$Macro'. Run: powershell -File build\macros.ps1"
        exit 6
    }
    $macroRoot = $mreg.macroRoot
    if ($env:VITTIX_MACRO_ROOT) { $macroRoot = $env:VITTIX_MACRO_ROOT }
    if (-not [System.IO.Path]::IsPathRooted($macroRoot)) { $macroRoot = Join-Path $repoRoot $macroRoot }
    $mdir     = Join-Path $macroRoot $macroEntry.dir
    $srcDir   = Join-Path $mdir 'src'
    $formsDir = Join-Path $mdir 'forms'
    if (-not $PSBoundParameters.ContainsKey('Project')) { $Project = "$($macroEntry.project).gms" }
    if ($macroEntry.PSObject.Properties.Name -contains 'importForms') { $importForms = [bool]$macroEntry.importForms }
}
elseif (-not $PSBoundParameters.ContainsKey('Project')) {
    $Project = 'vittix.gms'
}

if (-not $GmsDir) { $GmsDir = Get-GmsDir }

# ---------------------------------------------------------------- module set
if ($isMacro) {
    $modules = @(Get-ChildItem -LiteralPath $srcDir -File -ErrorAction SilentlyContinue |
                 Where-Object { $_.Extension -in '.bas', '.cls' })
    if ($importForms -and $formsDir -and (Test-Path -LiteralPath $formsDir)) {
        $modules += @(Get-ChildItem -LiteralPath $formsDir -File -ErrorAction SilentlyContinue |
                      Where-Object { $_.Extension -eq '.frm' })
    }
}
else {
    $modules = @(Get-ChildItem -LiteralPath $srcDir -File -ErrorAction SilentlyContinue |
                 Where-Object { $_.Extension -in '.bas', '.cls', '.frm' })
}
$modules = @($modules | Sort-Object Name)

Write-Host ("Target     : {0}{1}" -f $(if ($isMacro) { "macro '$Macro'" } else { 'framework project' }), "")
Write-Host ("Project    : {0}" -f $Project)
Write-Host ("GMS folder : {0}" -f $(if ($GmsDir) { $GmsDir } else { '<not found>' }))
Write-Host ("Source     : {0}" -f $srcDir)
Write-Host ("Modules    : {0}{1}" -f $modules.Count, $(if ($isMacro -and -not $importForms) { '  (.frm import disabled for this macro)' } else { '' }))

# ---------------------------------------------------------------- validation gate
if (-not $SkipValidate -and -not $List) {
    Write-Host "`n[1/4] Static validation ..."
    $va = @{}
    if ($Macro) { $va['Macro'] = $Macro }
    & (Join-Path $PSScriptRoot 'validate.ps1') @va
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Validation failed; aborting. (Use -SkipValidate to override.)"
        exit 5
    }
}

# ---------------------------------------------------------------- staging (bootstrap path)
if ($Stage) {
    $staging = Join-Path $env:APPDATA 'Vittix\staging'
    New-Item -ItemType Directory -Path $staging -Force | Out-Null
    $modules | ForEach-Object { Copy-Item -LiteralPath $_.FullName -Destination $staging -Force }
    Write-Host "`nStaged $($modules.Count) module(s) -> $staging"
    Write-Host "In CorelDRAW (target project active) run:  modDevImport.DevImport_RunAll"
}

# ---------------------------------------------------------------- connect
Write-Host "`n[2/4] Connecting to CorelDRAW ..."
$app = $null
$owned = $false
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
    Write-Host "  be attached to. Pass -StartCorel, or use the bootstrap path:"
    Write-Host "  build\deploy.ps1 -Macro $Macro -Stage  +  modDevImport.DevImport_RunAll"
    exit 2
}

$vbe = Get-VBE -App $app
if (-not $vbe) {
    Write-Host "CorelDRAW is running but Application.VBE is not accessible."
    Write-Host "Enable VBA project object-model access (Tools > Options > VBA in CorelDRAW),"
    Write-Host "then use the bootstrap path: build\deploy.ps1 -Macro $Macro -Stage  +  modDevImport.DevImport_RunAll"
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

# ---------------------------------------------------------------- find project
Write-Host "`n[3/4] Locating target project '$Project' ..."
$proj = $null
foreach ($p in $vbe.VBProjects) {
    $leaf = ''
    try { if ($p.FileName) { $leaf = Split-Path $p.FileName -Leaf } } catch { }
    if ($leaf -ieq $Project) { $proj = $p; break }
}
if (-not $proj) {
    Write-Host "Target project '$Project' is not loaded in this CorelDRAW instance."
    Write-Host "Load it first, or run -List to see what is loaded."
    exit 4
}

$dirty = $false
try { $dirty = -not $proj.Saved } catch { }
if ($dirty) { Write-Host "WARNING: target project has unsaved edits; this may interact with them." }

# ---------------------------------------------------------------- PULL
if ($Pull) {
    Write-Host "`nPull plan for $($proj.Name):"
    $pullPlan = @()
    foreach ($c in $proj.VBComponents) {
        $t = $c.Type
        $ext = $null
        if ($t -eq 1) { $ext = '.bas' } elseif ($t -eq 2) { $ext = '.cls' } elseif ($t -eq 3) { $ext = '.frm' }
        if ($ext) {
            $destDir = $srcDir
            if ($ext -eq '.frm') { $destDir = $(if ($formsDir) { $formsDir } else { $srcDir }) }
            $pullPlan += [pscustomobject]@{ Component = $c.Name; Kind = $ext; Dest = $destDir }
        }
    }
    $pullPlan | Format-Table -AutoSize | Out-String | Write-Host
    Write-Host "[4/4] " -NoNewline
    if (-not $Apply) { Write-Host "DRY RUN -- nothing written. Re-run with -Pull -Apply to export."; exit 0 }

    Write-Host "exporting ..."
    $ok = 0; $fail = 0
    foreach ($row in $pullPlan) {
        try {
            if (-not (Test-Path -LiteralPath $row.Dest)) { New-Item -ItemType Directory -Path $row.Dest -Force | Out-Null }
            $target = Join-Path $row.Dest ($row.Component + $row.Kind)
            foreach ($c in $proj.VBComponents) {
                if ($c.Name -ieq $row.Component) { $c.Export($target); break }
            }
            $ok++
            Write-Host ("  OK   " + $row.Component + $row.Kind)
        } catch {
            $fail++
            Write-Host ("  FAIL " + $row.Component + " :: " + $_.Exception.Message)
        }
    }
    Write-Host "`nExported $ok, failed $fail."
    if ($owned -and -not $KeepCorel) { try { $app.Quit() } catch { } }
    if ($fail -gt 0) { exit 1 }
    exit 0
}

# ---------------------------------------------------------------- PUSH plan
$plan = @()
foreach ($m in $modules) {
    $exists = $false
    foreach ($c in $proj.VBComponents) { if ($c.Name -ieq $m.BaseName) { $exists = $true; break } }
    $plan += [pscustomobject]@{
        Module = $m.BaseName
        Action = $(if ($exists) { 'Replace' } else { 'Add' })
        File   = $m.Name
    }
}
Write-Host "`nPush plan for $($proj.Name) ($($plan.Count) module(s)):"
$plan | Format-Table -AutoSize | Out-String | Write-Host

Write-Host "[4/4] " -NoNewline
if (-not $Apply) { Write-Host "DRY RUN -- nothing written. Re-run with -Apply to import."; exit 0 }

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

$saved = 'not attempted'
try {
    if ($proj.FileName) { $proj.SaveAs($proj.FileName); $saved = 'saved' }
    else { $saved = 'project has no FileName; save manually in the VBE' }
} catch { $saved = "save failed: $($_.Exception.Message)" }

Write-Host "`nImported $ok, failed $fail. Project save: $saved"
Write-Host "Note: a loaded .gms may also be persisted by CorelDRAW when the project is unloaded/closed - verify in the VBE."

if ($owned -and -not $KeepCorel) { try { $app.Quit() } catch { } }
if ($fail -gt 0) { exit 1 }
exit 0
