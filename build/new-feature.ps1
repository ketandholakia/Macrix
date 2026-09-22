<#
  build/new-feature.ps1 -- scaffold a new feature module that conforms to the
  project conventions (safe-run pattern, registry entry, docs/test stubs).

  Usage:
    powershell -File build/new-feature.ps1 -Name my-tool -Display "My Tool"
    powershell -File build/new-feature.ps1 -Name my-tool -Display "My Tool" -Stability Beta -Register

  Produces:
    src\feat_<PascalName>.bas   (from build\templates\feat_template.bas)
  and prints the exact registry line, a characterization-checklist row, and a
  CHANGELOG stub. With -Register it also inserts the registry line into
  modFeatureRegistry.FR_RegisterFeatures() for you.

  Exit code 0 = created; non-zero = refused/failed.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$Name,
    [Parameter(Mandatory = $true)][string]$Display,
    [ValidateSet('Stable', 'Beta', 'Experimental')][string]$Stability = 'Experimental',
    [string]$Version = '0.1.0',
    [string]$MinPlatform = '0.1.0',
    [switch]$Register,
    [switch]$Force
)

$ErrorActionPreference = 'Stop'

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$srcDir   = Join-Path $repoRoot 'src'
$tplPath  = Join-Path $PSScriptRoot 'templates\feat_template.bas'
$regPath  = Join-Path $srcDir 'modFeatureRegistry.bas'

if ($Name -notmatch '^[a-z][a-z0-9]*(-[a-z0-9]+)*$') {
    throw "Feature id '$Name' must be kebab-case (lowercase letters/digits, hyphen-separated), e.g. 'rounded-corners'."
}
if (-not (Test-Path $tplPath)) { throw "Template not found: $tplPath" }
if (-not (Test-Path $srcDir))  { throw "Source dir not found: $srcDir" }

# kebab -> PascalCase (rounded-corners -> RoundedCorners)
$pascal = ($Name -split '-' | Where-Object { $_.Length -gt 0 } | ForEach-Object {
    $_.Substring(0, 1).ToUpper() + $_.Substring(1)
}) -join ''
$module = "feat_$pascal"
$target = Join-Path $srcDir "$module.bas"

if ((Test-Path $target) -and -not $Force) {
    throw "$target already exists. Re-run with -Force to overwrite."
}

# Reserve an Err.Raise base above anything already used by features.
$maxErr = 0
Get-ChildItem -Path $srcDir -File -Filter 'feat_*.bas' -ErrorAction SilentlyContinue | ForEach-Object {
    $t = [System.IO.File]::ReadAllText($_.FullName, [System.Text.Encoding]::UTF8)
    foreach ($m in [regex]::Matches($t, 'Err\.Raise\s+(\d+)')) {
        $n = [int]$m.Groups[1].Value
        if ($n -gt $maxErr) { $maxErr = $n }
    }
}
$errBase = if ($maxErr -eq 0) { 1000 } else { ([math]::Floor($maxErr / 100) + 1) * 100 }

# Emit the module (UTF-8 without BOM, CRLF -- what VBE expects).
$body = [System.IO.File]::ReadAllText($tplPath, [System.Text.Encoding]::UTF8)
$body = $body.Replace('@@MODULE@@', $module).
              Replace('@@FEATURE_ID@@', $Name).
              Replace('@@DISPLAY@@', $Display).
              Replace('@@ERR_BASE@@', [string]$errBase)
$body = $body -replace "`r?`n", "`r`n"
[System.IO.File]::WriteAllText($target, $body, (New-Object System.Text.UTF8Encoding($false)))
Write-Host "Created $target  (Err base $errBase)"

# Build the registry line (printed, and optionally inserted).
$q = '"'
$fbStability = switch ($Stability) {
    'Stable'       { 'fb_Stable' }
    'Beta'         { 'fb_Beta' }
    'Experimental' { 'fb_Experimental' }
}
$regLine1 = '    FR_Register ' + $q + $Name + $q + ', ' + $q + $Display + $q + ', ' + $q + $module + $q + ', ' + $q + 'Main' + $q + ', _'
$regLine2 = '        fb_Available, ' + $fbStability + ', ' + $q + $Version + $q + ', ' + $q + $MinPlatform + $q

$registered = $false
if ($Register) {
    if (-not (Test-Path $regPath)) { throw "Registry file not found: $regPath" }
    $regText = [System.IO.File]::ReadAllText($regPath, [System.Text.Encoding]::UTF8)
    if ($regText -match [regex]::Escape($q + $module + $q)) {
        Write-Host "Already registered: $module is present in modFeatureRegistry (skipping insert)."
    }
    else {
        $srcLines = $regText -split "`r?`n"
        if ($srcLines.Count -gt 0 -and $srcLines[-1] -eq '') {
            $srcLines = $srcLines[0..($srcLines.Count - 2)]
        }
        $lines = [System.Collections.Generic.List[string]]::new()
        foreach ($ln in $srcLines) { [void]$lines.Add($ln) }

        $start = -1
        for ($i = 0; $i -lt $lines.Count; $i++) {
            if ($lines[$i] -match '^\s*Public\s+Sub\s+FR_RegisterFeatures\s*\(') { $start = $i; break }
        }
        if ($start -lt 0) { throw "FR_RegisterFeatures not found in $regPath" }

        $end = -1
        for ($i = $start + 1; $i -lt $lines.Count; $i++) {
            if ($lines[$i] -match '^\s*End\s+Sub\s*$') { $end = $i; break }
        }
        if ($end -lt 0) { throw "End Sub for FR_RegisterFeatures not found in $regPath" }

        $lines.Insert($end, $regLine2)
        $lines.Insert($end, $regLine1)
        $newText = ($lines -join "`r`n") + "`r`n"
        [System.IO.File]::WriteAllText($regPath, $newText, (New-Object System.Text.UTF8Encoding($false)))
        Write-Host "Registered '$Name' in modFeatureRegistry.FR_RegisterFeatures()."
        $registered = $true
    }
}

Write-Host ''
Write-Host 'Next steps'
Write-Host '----------'
Write-Host " 1. Implement p_DoWork in $module.bas (keep Main thin)."
if (-not $registered) {
    Write-Host ' 2. Paste this into modFeatureRegistry.FR_RegisterFeatures():'
    Write-Host $regLine1
    Write-Host $regLine2
}
else {
    Write-Host ' 2. Registry entry inserted.'
}
Write-Host " 3. tests\manual\CHARACTERIZATION.md -> add a '$Display' row (precondition, steps, expected)."
Write-Host ' 4. CHANGELOG.md -> under [Unreleased] > Added, one line for this feature.'
Write-Host ' 5. powershell -File build\validate.ps1        # must report 0 errors'
Write-Host ' 6. modTestRunner.RunAll + manual characterization on a real CorelDRAW install.'
