<#
* build/validate.ps1 — static lint / validation.
*
* Runs in CI *without* CorelDRAW (PowerShell 5+ / pwsh on Windows). Per §41 this
* is the ONLY step that is automated in CI; COM behavior is manual characterization.
*
* Checks:
*   1. Option Explicit present in every .bas/.cls.
*   2. 'Attribute VB_Name' present and matching the file's base name (.bas/.cls).
*   3. Balanced procedure blocks (Sub/Function/Property vs End ...).
*   4. Duplicate public procedure detection within a module.
*   5. Public-proc naming (should start uppercase; namespace prefixes recognized).
*   6. Leading-underscore identifiers — invalid VBA, a historical project bug.
*   7. Hard-coded absolute Windows paths.
*   8. Feature registry drift:
*        - every src\feat_*.bas is registered in modFeatureRegistry, and
*        - every module named by an FR_Register call actually exists.
*
* Exit code 0 = OK, 1 = errors, 2 = warnings-only.
#>
param(
    [string]$SrcDir = (Join-Path $PSScriptRoot '..\src')
)
$ErrorActionPreference = 'Continue'

$errors   = [System.Collections.ArrayList]::new()
$warnings = [System.Collections.ArrayList]::new()
$signatures = @{}   # "Module.Proc" -> $true

function AddError([string]$msg) { [void]$errors.Add($msg) }
function AddWarn ([string]$msg) { [void]$warnings.Add($msg) }

if (-not (Test-Path $SrcDir)) {
    AddError "Source dir not found: $SrcDir"
    Write-Host 'ERROR: no src dir'
    exit 1
}

Write-Host "Scanning $SrcDir ..."

$files = Get-ChildItem -Path $SrcDir -Recurse -File -Include *.bas, *.cls, *.frm -ErrorAction SilentlyContinue
if (-not $files) { AddWarn "No .bas/.cls/.frm found; nothing to validate." }

foreach ($file in $files) {
    $name = $file.Name
    $text = Get-Content -Raw -LiteralPath $file.FullName
    if ($null -eq $text) { continue }
    $lines = $text -split "`r?`n"

    # 1. Option Explicit
    if ($file.Extension -ne '.frm' -and $text -notmatch '(?im)^\s*Option\s+Explicit\s*$') {
        AddError "$name :: missing Option Explicit"
    }

    # 2. Attribute VB_Name matches file name (.bas/.cls)
    if ($file.Extension -ne '.frm') {
        $vm = [regex]::Match($text, '(?im)^\s*Attribute\s+VB_Name\s*=\s*"([^"]+)"')
        if (-not $vm.Success) {
            AddError "$name :: missing 'Attribute VB_Name'"
        }
        elseif ($vm.Groups[1].Value -ne $file.BaseName) {
            AddError "$name :: Attribute VB_Name '$($vm.Groups[1].Value)' != file base name '$($file.BaseName)'"
        }
    }

    # 3-7. per-line checks
    $decls = 0
    $ends  = 0
    foreach ($line in $lines) {
        $t = $line -replace '^\s+', ''

        if ($t -match '^[A-Za-z]:\\') { AddWarn "$name :: hard-coded path -> $line".Trim() }

        # 6. leading-underscore identifier (uncompilable in VBA)
        if ($t -match '^\s*(Public\s+|Private\s+|Friend\s+)?(Sub|Function|Property\s+(Get|Let|Set))\s+(_[A-Za-z0-9_]*)') {
            AddError "$name :: identifier starting with '_' is not valid VBA ($t)".Trim()
        }

        # 4/5. public Sub/Function
        if ($t -match '^\s*(Public\s+|Private\s+|Friend\s+)?(Sub|Function)\s+([A-Za-z][A-Za-z0-9_]*)\s*\(') {
            $decls++
            $proc  = $Matches[3]
            $scope = if ($t -match '^\s*(Private|Friend)\b') { 'private' } else { 'public' }
            if ($scope -eq 'public') {
                $key = "$($file.BaseName).$proc"
                if ($signatures.ContainsKey($key)) { AddError "duplicate public procedure: $key" }
                else { $signatures[$key] = $true }
                if ($proc -notmatch '^[A-Z]') {
                    AddWarn "$name :: public proc '$proc' should start uppercase"
                }
            }
        }

        # 4/5. public Property Get/Let/Set
        if ($t -match '^\s*(Public\s+|Private\s+|Friend\s+)?Property\s+(Get|Let|Set)\s+([A-Za-z][A-Za-z0-9_]*)\s*\(') {
            $decls++
            $proc  = $Matches[3]
            $scope = if ($t -match '^\s*(Private|Friend)\b') { 'private' } else { 'public' }
            if ($scope -eq 'public') {
                $key = "$($file.BaseName).$proc"
                if ($signatures.ContainsKey($key)) { AddError "duplicate public procedure: $key" }
                else { $signatures[$key] = $true }
                if ($proc -notmatch '^[A-Z]') {
                    AddWarn "$name :: public proc '$proc' should start uppercase"
                }
            }
        }

        # 3. procedure terminators
        if ($t -match '^\s*End\s+(Sub|Function|Property)\b') { $ends++ }
    }

    if ($file.Extension -ne '.frm' -and $decls -ne $ends) {
        AddError "$name :: unbalanced procedure blocks (decls=$decls, ends=$ends)"
    }
}

# 8. Feature registry <-> module drift.
$regFile = Join-Path $SrcDir 'modFeatureRegistry.bas'
if (Test-Path $regFile) {
    $regText = Get-Content -Raw -LiteralPath $regFile
    $registered = @{}   # module name -> feature id
    foreach ($m in [regex]::Matches($regText, 'FR_Register\s+"([^"]+)"\s*,\s*"([^"]+)"\s*,\s*"([^"]+)"')) {
        $registered[$m.Groups[3].Value] = $m.Groups[1].Value
    }

    foreach ($mod in $registered.Keys) {
        if (-not (Test-Path (Join-Path $SrcDir "$mod.bas"))) {
            AddError "modFeatureRegistry registers module '$mod' but src\$mod.bas does not exist"
        }
    }

    foreach ($f in (Get-ChildItem -Path $SrcDir -File -Filter 'feat_*.bas' -ErrorAction SilentlyContinue)) {
        if (-not $registered.ContainsKey($f.BaseName)) {
            AddError "$($f.Name) :: feature module not registered in modFeatureRegistry.FR_RegisterFeatures()"
        }
    }

    if ($registered.Count -eq 0) { AddWarn 'no features registered in modFeatureRegistry' }
}
else {
    AddWarn 'modFeatureRegistry.bas not found; skipping registry drift checks'
}

Write-Host ("errors: {0}  warnings: {1}" -f $errors.Count, $warnings.Count)
$errors   | ForEach-Object { Write-Host ('ERROR  ' + $_) }
$warnings | ForEach-Object { Write-Host ('WARN   ' + $_) }
if ($errors.Count -gt 0) { exit 1 }
if ($warnings.Count -gt 0) { exit 2 }
exit 0
