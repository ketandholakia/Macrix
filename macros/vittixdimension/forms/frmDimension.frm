VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDimension 
   Caption         =   "Vittix Dimension"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7170
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
    LoadFormState
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
        cmbPosition.AddItem "Above"
        cmbPosition.AddItem "Below"
        cmbPosition.AddItem "Left"
        cmbPosition.AddItem "Right"
    End If
End Sub

