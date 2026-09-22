<#
  macros/vittixdimension/scripts/Build-Form.ps1

  Builds the frmDimension UserForm as a REAL design-time form inside the
  VittixDimensionTools VBA project, then lets CorelDRAW persist it to the .gms.

  WHY THIS EXISTS
  The .frm/.frx exported from this project is not importable in a usable way:
  its control definitions live in the binary .frx, and importing the .frm into
  CorelDRAW's VBE yields a form with NO controls. Any code referencing controls
  (cmbUnit, txtTextWidthPercent, ...) then fails to compile with
  "Method or data member not found" (461). This script instead builds the form
  through the VBIDE Designer API, where the controls are real.

  USAGE
    powershell -File macros\vittixdimension\scripts\Build-Form.ps1

  NOTES
  - Requires CorelDRAW installed (launches a headless /Automation instance and
    quits it). Close any interactive CorelDRAW first.
  - After running, deploy.ps1 -Macro dimension-tools -Pull -Apply re-exports the
    form back into forms\frmDimension.frm|.frx.
  - Keep the registry's "importForms": false for this macro: importing the .frm
    would reintroduce the broken, control-less form.
#>
$ErrorActionPreference = 'Continue'
$gmsProject = 'VittixDimensionTools'

$app = $null
try {
    $app = New-Object -ComObject CorelDRAW.Application
    try { $app.Visible = $false } catch {}
    try { [void]$app.InitializeVBA() } catch {}
    Start-Sleep -Seconds 2
    $vbe = $app.VBE

    $proj = $null
    foreach ($p in $vbe.VBProjects) {
        $fn = ''; try { $fn = $p.FileName } catch {}
        if ($fn -and (Split-Path $fn -Leaf) -ieq "$gmsProject.gms") { $proj = $p; break }
    }
    if (-not $proj) { throw "project '$gmsProject.gms' not found" }

    foreach ($c in @($proj.VBComponents)) {
        if ($c.Name -ieq 'frmDimension') { $proj.VBComponents.Remove($c); Write-Host 'removed existing frmDimension' }
    }

    $form = $proj.VBComponents.Add(3)   # 3 = vbext_ct_MSForm
    $form.Name = 'frmDimension'
    Write-Host ('created form component: ' + $form.Name)

    $d = $form.Designer
    function AddCtl($progId, $name, $left, $top, $w, $h) {
        $c = $d.Controls.Add($progId, $name)
        $c.Left = $left; $c.Top = $top; $c.Width = $w; $c.Height = $h
        return $c
    }
    function Lbl($n, $cap, $l, $t, $w)     { $c = AddCtl 'Forms.Label.1' $n $l $t $w 17; $c.Caption = $cap }
    function Cmb($n, $l, $t)               { $c = AddCtl 'Forms.ComboBox.1' $n $l $t 160 20; $c.Style = 2 }
    function Txt($n, $l, $t, $w)           { AddCtl 'Forms.TextBox.1' $n $l $t $w 20 | Out-Null }
    function Chk($n, $cap, $l, $t, $w, $v) { $c = AddCtl 'Forms.CheckBox.1' $n $l $t $w 17; $c.Caption = $cap; $c.Value = $v }
    function Btn($n, $cap, $l, $t)         { $c = AddCtl 'Forms.CommandButton.1' $n $l $t 90 25; $c.Caption = $cap }
    function Swt($n, $l, $t, $w)           { $c = AddCtl 'Forms.Label.1' $n $l $t $w 16; $c.Caption = ''; $c.BackColor = 0; $c.BorderStyle = 1 }

    Lbl 'lblUnit' 'Unit' 16 16 100;                      Cmb 'cmbUnit' 120 13
    Lbl 'lblDecimals' 'Decimals' 16 46 100;              Cmb 'cmbDecimals' 120 43
    Lbl 'lblPosition' 'Position' 16 76 100;              Cmb 'cmbPosition' 120 73
    Lbl 'lblTextColor' 'Text Color' 16 106 100;          Txt 'txtTextColor' 120 106 130; Swt 'lblColorSwatch' 256 107 40
    Lbl 'lblTextWidthPercent' 'Text Width %' 16 136 100; Txt 'txtTextWidthPercent' 120 136 160
    Lbl 'lblGap' 'Gap (mm)' 16 166 100;                  Txt 'txtGap' 120 166 160
    Lbl 'lblPadding' 'Padding (mm)' 16 196 100;          Txt 'txtPadding' 120 196 160
    Lbl 'lblCornerRadius' 'Corner Radius (mm)' 16 226 130; Txt 'txtCornerRadius' 150 226 130
    Chk 'chkShowWidth' 'Show Width' 16 261 130 $true
    Chk 'chkShowHeight' 'Show Height' 16 283 130 $true
    Chk 'chkShowArea' 'Show Area' 16 305 130 $false
    Chk 'chkShowPerimeter' 'Show Perimeter' 16 327 130 $false
    Chk 'chkShowObjectCount' 'Show Object Count' 150 261 150 $false
    Chk 'chkBackgroundBox' 'Background Box' 150 283 150 $true
    Chk 'chkRoundedBackground' 'Rounded Background' 150 305 150 $false
    Chk 'chkCreateLayer' 'Create Layer' 150 327 150 $true
    Chk 'chkRememberSettings' 'Remember Settings' 16 351 200 $true
    Lbl 'lblTemplate' 'Template' 16 381 100;             Txt 'txtTemplate' 120 381 160
    Btn 'cmdOK' 'OK' 60 416
    Btn 'cmdCancel' 'Cancel' 160 416

    # OK is the default button (Enter); Cancel responds to Esc.
    $d.Controls.Item('cmdOK').Default = $true
    $d.Controls.Item('cmdCancel').Cancel = $true
    Write-Host 'cmdOK set as default, cmdCancel as cancel'
    Write-Host ('controls added: ' + $d.Controls.Count)

    # Size the form to contain every control. Note: the form property setters take
    # STRING values, and Width/Height are POINTS covering the border, while control
    # Left/Top are inner coordinates -- hence the margins. Without this the form keeps
    # the default 240x180 and clips everything below the top rows.
    $maxR = 0; $maxB = 0
    foreach ($ct in $d.Controls) {
        $r = [double]$ct.Left + [double]$ct.Width
        $b = [double]$ct.Top + [double]$ct.Height
        if ($r -gt $maxR) { $maxR = $r }
        if ($b -gt $maxB) { $maxB = $b }
    }
    $form.Properties.Item('Caption').Value = 'Vittix Dimension'
    $form.Properties.Item('Width').Value = [string]([int]([math]::Ceiling($maxR) + 30))
    $form.Properties.Item('Height').Value = [string]([int]([math]::Ceiling($maxB) + 45))
    $form.Properties.Item('StartUpPosition').Value = '1'
    Write-Host ('form size set to {0} x {1} pt (controls need {2} x {3})' -f $form.Properties.Item('Width').Value, $form.Properties.Item('Height').Value, $maxR, $maxB)

    $code = @'
Option Explicit

Private Sub UserForm_Initialize()
    LoadFormState
    On Error Resume Next
    ' Enter activates OK, Esc activates Cancel.
    cmdOK.Default = True
    cmdCancel.Cancel = True
End Sub

' Live preview of the caption colour as it is typed (#RRGGBB or R,G,B).
Private Sub txtTextColor_Change()
    On Error Resume Next
    Dim c As Long
    c = ColorFromHex(txtTextColor.text)
    If c >= 0 Then lblColorSwatch.BackColor = c
End Sub

Public Sub LoadFormState()
    On Error Resume Next
    PopulateControls
    EnsureFormDefaults
    ApplySettingsToForm Me
End Sub

Private Sub cmdOK_Click()
    On Error GoTo ErrHandler
    ReadSettingsFromForm Me, gSettings
    SaveSettings
    VittixDimensionTools_Create
    Unload Me
    Exit Sub
ErrHandler:
    LogError "frmDimension.cmdOK_Click", Err.Number, Err.Description
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Unload Me
    End If
End Sub

Private Sub EnsureFormDefaults()
    On Error Resume Next
    If Len(CStr(cmbUnit.value)) = 0 Then cmbUnit.value = "MM"
    If Len(CStr(cmbDecimals.value)) = 0 Then cmbDecimals.value = "2"
    If Len(CStr(cmbPosition.value)) = 0 Then cmbPosition.value = "Auto"
End Sub

Private Sub PopulateControls()
    On Error Resume Next
    If cmbUnit.ListCount = 0 Then
        cmbUnit.AddItem "MM"
        cmbUnit.AddItem "CM"
        cmbUnit.AddItem "M"
        cmbUnit.AddItem "IN"
        cmbUnit.AddItem "FT"
    End If
    If cmbDecimals.ListCount = 0 Then
        cmbDecimals.AddItem "0"
        cmbDecimals.AddItem "1"
        cmbDecimals.AddItem "2"
        cmbDecimals.AddItem "3"
        cmbDecimals.AddItem "4"
    End If
    If cmbPosition.ListCount = 0 Then
        cmbPosition.AddItem "Auto"
        cmbPosition.AddItem "Center"
        cmbPosition.AddItem "Above"
        cmbPosition.AddItem "Below"
        cmbPosition.AddItem "Left"
        cmbPosition.AddItem "Right"
    End If
End Sub
'@
    $cm = $form.CodeModule
    if ($cm.CountOfLines -gt 0) { $cm.DeleteLines(1, $cm.CountOfLines) }
    $cm.AddFromString($code)
    Write-Host ('code-behind lines: ' + $form.CodeModule.CountOfLines)

    # Export the rebuilt form back into the repo so source tracks reality.
    $formsDir = Join-Path (Split-Path $PSScriptRoot -Parent) 'forms'
    if (Test-Path -LiteralPath $formsDir) {
        $frmPath = Join-Path $formsDir 'frmDimension.frm'
        $form.Export($frmPath)
        Write-Host ('exported form -> ' + $frmPath)
    }
    Write-Host 'form build done. CorelDRAW will persist it to the .gms on exit.'
}
catch {
    Write-Host ('ERROR: ' + $_.Exception.Message)
    Write-Host 'NOTE: removing an existing form and adding a new one in the SAME session can fail'
    Write-Host '      with a path/file access error. If that happens, re-run this script: the old'
    Write-Host '      form is already gone, so the plain rebuild (add) then succeeds.'
}
finally {
    if ($app) { try { $app.Quit() } catch {} }
    Write-Host 'done.'
}
