<#
* build/validate.ps1 -- static lint / validation.
*
* Runs in CI *without* CorelDRAW (PowerShell 5+ / pwsh on Windows). Per sec.41 this
* is the ONLY step that is automated in CI; COM behavior is manual characterization.
*
* Checks:
*   1. Option Explicit present in every .bas/.cls.
*   2. 'Attribute VB_Name' present and matching the file's base name (.bas/.cls).
*   3. Balanced procedure blocks (Sub/Function/Property vs End ...).
*   4. Duplicate public procedure detection within a module.
*   5. Public-proc naming (should start uppercase; namespace prefixes recognized).
*   6. Leading-underscore identifiers -- invalid VBA, a historical project bug.
*   6b. Optional parameter typed as a user-defined Type -- invalid VBA
*       ("Invalid optional parameter type").
*   7. Hard-coded absolute Windows paths.
*   8. Feature registry drift:
*        - every src\feat_*.bas is registered in modFeatureRegistry, and
*        - every module named by an FR_Register call actually exists.
*   9. Macro registry (macros\registry.json): each registered macro's folder and
*      entry-point modules exist in the sibling macro checkout. Warns (does not
*      fail) when the macro root is not present, so CI stays green off-machine.
*  10. Optional -Macro <id>: additionally lint that macro's src\ (structural
*      checks only; host document modules are exempt from Option Explicit).
*
* Structural checks (1-7) also cover tools\dev-import; checks 8-9 are src/macros.
*
* Exit code 0 = OK, 1 = errors, 2 = warnings-only.
#>
param(
    [string]$SrcDir = (Join-Path $PSScriptRoot '..\src'),
    [string]$Macro
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

# Structural checks run on src\ plus the dev-tooling folder; registry drift
# (check 8) applies to src\ only.
$scanDirs = @($SrcDir)
$toolsDir = Join-Path (Split-Path $SrcDir -Parent) 'tools\dev-import'
if (Test-Path $toolsDir) { $scanDirs += $toolsDir }

# Optional (-Macro <id>): additionally lint one registered macro's sources, so a
# problem in one legacy macro cannot block work on another.
if ($Macro) {
    $repoRootV = Split-Path $PSScriptRoot -Parent
    $regFileV  = Join-Path $repoRootV 'macros\registry.json'
    if (-not (Test-Path -LiteralPath $regFileV)) {
        AddError "unknown macro '$Macro': macro registry not found"
    }
    else {
        try {
            $mregV = Get-Content -Raw -LiteralPath $regFileV -Encoding UTF8 | ConvertFrom-Json
            $mentry = $mregV.macros | Where-Object { $_.id -eq $Macro }
            if (-not $mentry) {
                AddError "unknown macro id '$Macro'"
            }
            else {
                $macroRootV = $mregV.macroRoot
                if ($env:VITTIX_MACRO_ROOT) { $macroRootV = $env:VITTIX_MACRO_ROOT }
                if (-not [System.IO.Path]::IsPathRooted($macroRootV)) { $macroRootV = Join-Path $repoRootV $macroRootV }
                $macroSrcV = Join-Path (Join-Path $macroRootV $mentry.dir) 'src'
                if (Test-Path -LiteralPath $macroSrcV) {
                    Write-Host "macro lint: $Macro ($($mentry.project))"
                    $scanDirs += $macroSrcV
                }
                else { AddError "macro '$Macro': src dir not found ($macroSrcV)" }
            }
        }
        catch { AddError "macro registry could not be parsed: $($_.Exception.Message)" }
    }
}

$files = foreach ($d in $scanDirs) {
    Get-ChildItem -Path $d -Recurse -File -Include *.bas, *.cls, *.frm -ErrorAction SilentlyContinue
}
if (-not $files) { AddWarn "No .bas/.cls/.frm found; nothing to validate." }

# Collect user-defined Type names (used by the Optional-parameter rule below).
$udtNames = @{}
foreach ($f in @($files)) {
    $ft = Get-Content -Raw -LiteralPath $f.FullName
    if ($ft) {
        foreach ($m in [regex]::Matches($ft, '(?im)^\s*(?:Public\s+|Private\s+)?Type\s+([A-Za-z][A-Za-z0-9_]*)')) {
            $udtNames[$m.Groups[1].Value] = $true
        }
    }
}

foreach ($file in $files) {
    $name = $file.Name
    $text = Get-Content -Raw -LiteralPath $file.FullName
    if ($null -eq $text) { continue }
    $lines = $text -split "`r?`n"

    # 1. Option Explicit (host document modules never declare it)
    $isDocModule = $file.BaseName -match '^(ThisMacroStorage|ThisDocument|ThisWorkbook|ThisDrawing)$'
    if ($file.Extension -ne '.frm' -and -not $isDocModule -and $text -notmatch '(?im)^\s*Option\s+Explicit\s*$') {
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

        # 6b. VBA forbids an Optional parameter whose type is a user-defined Type
        # ("Invalid optional parameter type"). This bit the dimension macro once.
        if ($t -match '(?i)\bOptional\s+[A-Za-z][A-Za-z0-9_]*\s+As\s+([A-Za-z][A-Za-z0-9_]*)') {
            $optType = $Matches[1]
            if ($udtNames.ContainsKey($optType)) {
                AddError "$name :: 'Optional' parameter cannot be a user-defined type ('$optType')"
            }
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

# 9. Macro registry <-> sibling macro checkout drift.
$repoRootForMacros = Split-Path $PSScriptRoot -Parent
$macroRegFile = Join-Path $repoRootForMacros 'macros\registry.json'
if (Test-Path -LiteralPath $macroRegFile) {
    try {
        $mreg = Get-Content -Raw -LiteralPath $macroRegFile -Encoding UTF8 | ConvertFrom-Json
        $macroRoot = $mreg.macroRoot
        if ($env:VITTIX_MACRO_ROOT) { $macroRoot = $env:VITTIX_MACRO_ROOT }
        if (-not [System.IO.Path]::IsPathRooted($macroRoot)) {
            $macroRoot = Join-Path $repoRootForMacros $macroRoot
        }
        if (-not (Test-Path -LiteralPath $macroRoot)) {
            # Informational, not a warning: the macro checkout is expected to be
            # absent on CI, and a warning would flip the exit code to 2 and fail CI.
            Write-Host "macro registry: macro root not found ($macroRoot); skipping macro checks"
        }
        else {
            foreach ($mac in $mreg.macros) {
                $mdir = Join-Path $macroRoot $mac.dir
                if (-not (Test-Path -LiteralPath $mdir)) {
                    AddError "macro registry: '$($mac.id)' folder missing ($($mac.dir))"
                    continue
                }
                foreach ($ep in $mac.entryPoints) {
                    $parts = $ep -split '\.'
                    $mod   = $parts[0]
                    $proc  = if ($parts.Count -gt 1) { $parts[1] } else { $null }
                    $modFile = $null
                    foreach ($ext in '.bas', '.cls') {
                        $cand = Join-Path $mdir "src\$mod$ext"
                        if (Test-Path -LiteralPath $cand) { $modFile = $cand; break }
                    }
                    if (-not $modFile) {
                        AddError "macro registry: '$($mac.id)' entry point '$ep' -> module '$mod' not found"
                        continue
                    }
                    if ($proc) {
                        $mtext = Get-Content -Raw -LiteralPath $modFile
                        if ($mtext -notmatch "(?im)^\s*(Public\s+|Private\s+|Friend\s+)?(Sub|Function)\s+$([regex]::Escape($proc))\s*\(") {
                            AddError "macro registry: '$($mac.id)' entry point '$ep' -> procedure '$proc' not found in $mod"
                        }
                    }
                }
            }
        }
    }
    catch { AddWarn "macro registry could not be parsed: $($_.Exception.Message)" }
}

Write-Host ("errors: {0}  warnings: {1}" -f $errors.Count, $warnings.Count)
$errors   | ForEach-Object { Write-Host ('ERROR  ' + $_) }
$warnings | ForEach-Object { Write-Host ('WARN   ' + $_) }
if ($errors.Count -gt 0) { exit 1 }
if ($warnings.Count -gt 0) { exit 2 }
exit 0
