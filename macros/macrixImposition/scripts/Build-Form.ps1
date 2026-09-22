<#
  macros/macrixImposition/scripts/Build-Form.ps1

  Builds frmImpositionSettings as a REAL design-time form inside the
  macrixImposition VBA project, then lets CorelDRAW persist it to the .gms.

  WHY THIS EXISTS
  The checked-in .frm was VB6 format (Begin VB.Form / LinkTopic), which CorelDRAW's VBE
  cannot import. The original workaround built the form at run time from mdlFormBuilder,
  but that path is fragile: VBComponents.Add can return Nothing, and the very next line
  (newFormComp.Name = ...) then raises "Run-time error 424: Object required".
  Building the form once, design-time, removes the runtime VBE dependency entirely.

  The control list and the code-behind are EXTRACTED from src\mdlFormBuilder.bas, so this
  script stays in step with that module instead of duplicating it.

  Usage:  powershell -File macros\macrixImposition\scripts\Build-Form.ps1
#>
$ErrorActionPreference = 'Stop'

$live = Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.ProcessName -match '(?i)corel' }
if ($live) { Write-Host 'CorelDRAW is running - close it first.'; exit 1 }

$macroDir = Split-Path $PSScriptRoot -Parent
$srcPath  = Join-Path $macroDir 'src\mdlFormBuilder.bas'
$frmOut   = Join-Path $macroDir 'forms\frmImpositionSettings.frm'
$formName = 'frmImpositionSettings'

$src = [System.IO.File]::ReadAllText($srcPath)
$lines = $src -split "`r?`n"

$ctlCalls = @()
foreach ($ln in $lines) {
    if ($ln -match '^\s*Add(Label|ComboList|TextBox|CheckBox|Button)\s+d,\s*(.+?)\s*$') {
        $ctlCalls += [pscustomobject]@{ Kind = $Matches[1]; Args = $Matches[2] }
    }
}
Write-Host ('controls found: ' + $ctlCalls.Count)

$sb = New-Object System.Text.StringBuilder
$inInject = $false
foreach ($ln in $lines) {
    if ($ln -match 'Sub\s+InjectCodeBehind') { $inInject = $true; continue }
    if (-not $inInject) { continue }
    if ($ln -match 'codeMod\.DeleteLines') { break }
    if ($ln -match 's\s*=\s*s\s*&\s*"(.*)"\s*&\s*vbCrLf\s*$') {
        [void]$sb.AppendLine($Matches[1].Replace('""', '"'))
    }
}
$code = $sb.ToString()
Write-Host ('injected code lines: ' + ($code -split "`r?`n").Count)

function Split-Args([string]$raw) {
    $out = @()
    foreach ($p in ($raw -split ',')) {
        $v = $p.Trim()
        if ($v.StartsWith('"') -and $v.EndsWith('"')) { $v = $v.Substring(1, $v.Length - 2) }
        $out += $v
    }
    return $out
}

$app = $null
try {
    for ($try = 1; $try -le 3; $try++) {
        try { $app = New-Object -ComObject CorelDRAW.Application; break } catch { Start-Sleep -Seconds 8 }
    }
    if (-not $app) { throw 'could not launch CorelDRAW' }
    try { $app.Visible = $false } catch {}
    try { [void]$app.InitializeVBA() } catch {}
    Start-Sleep -Seconds 5

    $vbe = $app.VBE
    $proj = $null
    foreach ($p in $vbe.VBProjects) {
        $fn = ''; try { $fn = $p.FileName } catch {}
        if ($fn -and (Split-Path $fn -Leaf) -ieq 'macrixImposition.gms') { $proj = $p; break }
    }
    if (-not $proj) { throw 'macrixImposition.gms not found' }
    Write-Host ('project: ' + $proj.Name)

    foreach ($c in @($proj.VBComponents)) {
        if ($c.Name -ieq $formName) { $proj.VBComponents.Remove($c); Write-Host 'removed existing form' }
    }

    $form = $proj.VBComponents.Add(3)
    if ($form -eq $null) { throw 'VBComponents.Add(3) returned $null' }
    $form.Name = $formName
    Write-Host 'form component created'

    $form.Properties.Item('Caption').Value = 'Imposition Settings'
    $form.Properties.Item('Width').Value = [string]310
    $form.Properties.Item('Height').Value = [string]452

    $d = $form.Designer
    foreach ($call in $ctlCalls) {
        $a = Split-Args $call.Args
        $name = $a[0]
        switch ($call.Kind) {
            'Label'     { $c = $d.Controls.Add('Forms.Label.1', $name);         $c.Caption = $a[1]; $c.Left = [single]$a[2]; $c.Top = [single]$a[3]; $c.Width = [single]$a[4]; $c.Height = [single]$a[5] }
            'ComboList' { $c = $d.Controls.Add('Forms.ComboBox.1', $name);      $c.Left = [single]$a[1]; $c.Top = [single]$a[2]; $c.Width = [single]$a[3]; $c.Height = [single]$a[4]; $c.Style = 2 }
            'TextBox'   { $c = $d.Controls.Add('Forms.TextBox.1', $name);       $c.Text = $a[1]; $c.Left = [single]$a[2]; $c.Top = [single]$a[3]; $c.Width = [single]$a[4]; $c.Height = [single]$a[5] }
            'CheckBox'  { $c = $d.Controls.Add('Forms.CheckBox.1', $name);      $c.Caption = $a[1]; $c.Left = [single]$a[2]; $c.Top = [single]$a[3]; $c.Width = [single]$a[4]; $c.Value = [bool]::Parse($a[5]) }
            'Button'    { $c = $d.Controls.Add('Forms.CommandButton.1', $name); $c.Caption = $a[1]; $c.Left = [single]$a[2]; $c.Top = [single]$a[3]; $c.Width = [single]$a[4]; $c.Height = [single]$a[5] }
        }
    }
    Write-Host ('controls added: ' + $d.Controls.Count)

    $cm = $form.CodeModule
    if ($cm.CountOfLines -gt 0) { $cm.DeleteLines(1, $cm.CountOfLines) }
    $cm.AddFromString($code)
    Write-Host ('code-behind lines: ' + $form.CodeModule.CountOfLines)

    $form.Export($frmOut)
    Write-Host ('exported -> ' + $frmOut)
}
catch { Write-Host ('ERROR: ' + $_.Exception.Message) }
finally {
    if ($app) { try { $app.Quit() } catch {} }
    Write-Host 'done.'
}
