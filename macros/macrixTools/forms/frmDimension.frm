VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDimension 
   Caption         =   "Macrix Dimension"
   ClientHeight    =   10365
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6510
   OleObjectBlob   =   "frmDimension.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDimension"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Me.Caption = "Macrix Dimension"
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
    If ActiveDocument Is Nothing Then
        MsgBox "Open a document before creating dimensions.", vbExclamation, "Vittix"
        Exit Sub
    End If
    CreateDimensionsForSelection gSettings
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

Private Sub SetColorFromSwatch(ByVal ctlName As String)
    On Error Resume Next
    Dim ctl As Object
    Set ctl = Me.Controls(ctlName)
    If ctl Is Nothing Then Exit Sub
    txtTextColor.text = CStr(ctl.Tag)
End Sub

Private Sub sw01_Click()
    SetColorFromSwatch "sw01"
End Sub

Private Sub sw02_Click()
    SetColorFromSwatch "sw02"
End Sub

Private Sub sw03_Click()
    SetColorFromSwatch "sw03"
End Sub

Private Sub sw04_Click()
    SetColorFromSwatch "sw04"
End Sub

Private Sub sw05_Click()
    SetColorFromSwatch "sw05"
End Sub

Private Sub sw06_Click()
    SetColorFromSwatch "sw06"
End Sub

Private Sub sw07_Click()
    SetColorFromSwatch "sw07"
End Sub

Private Sub sw08_Click()
    SetColorFromSwatch "sw08"
End Sub

Private Sub sw09_Click()
    SetColorFromSwatch "sw09"
End Sub

Private Sub sw10_Click()
    SetColorFromSwatch "sw10"
End Sub

Private Sub sw11_Click()
    SetColorFromSwatch "sw11"
End Sub

Private Sub sw12_Click()
    SetColorFromSwatch "sw12"
End Sub

Private Sub sw13_Click()
    SetColorFromSwatch "sw13"
End Sub

Private Sub sw14_Click()
    SetColorFromSwatch "sw14"
End Sub

Private Sub sw15_Click()
    SetColorFromSwatch "sw15"
End Sub

Private Sub sw16_Click()
    SetColorFromSwatch "sw16"
End Sub

Private Sub sw17_Click()
    SetColorFromSwatch "sw17"
End Sub

Private Sub sw18_Click()
    SetColorFromSwatch "sw18"
End Sub

Private Sub sw19_Click()
    SetColorFromSwatch "sw19"
End Sub

Private Sub sw20_Click()
    SetColorFromSwatch "sw20"
End Sub

Private Sub sw21_Click()
    SetColorFromSwatch "sw21"
End Sub

Private Sub sw22_Click()
    SetColorFromSwatch "sw22"
End Sub

Private Sub sw23_Click()
    SetColorFromSwatch "sw23"
End Sub

Private Sub sw24_Click()
    SetColorFromSwatch "sw24"
End Sub

