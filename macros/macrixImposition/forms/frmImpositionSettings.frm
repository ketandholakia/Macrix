VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmImpositionSettings 
   Caption         =   "Imposition Settings"
   ClientHeight    =   8610.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6105
   OleObjectBlob   =   "frmImpositionSettings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmImpositionSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()

    If Not g_SettingsInitialized Then InitDefaultSettings

    cboLayoutMode.AddItem "Simple"
    cboLayoutMode.AddItem "StepRepeat"
    cboLayoutMode.AddItem "Signature"
    cboLayoutMode.ListIndex = IndexOfString(cboLayoutMode, g_LayoutMode)

    cboUnit.AddItem "Millimeters"
    cboUnit.AddItem "Inches"
    cboUnit.AddItem "Points"
    cboUnit.AddItem "Pixels"
    cboUnit.ListIndex = UnitToComboIndex(g_Unit)

    txtSheetWidth.Text = g_SheetWidth
    txtSheetHeight.Text = g_SheetHeight
    txtGridRows.Text = g_GridRows
    txtGridCols.Text = g_GridCols
    txtGutterX.Text = g_GutterX
    txtGutterY.Text = g_GutterY
    txtMarginLeft.Text = g_MarginLeft
    txtMarginTop.Text = g_MarginTop
    txtBleedSize.Text = g_BleedSize
    chkCropMarks.Value = IIf(g_AddCropMarks, 1, 0)
    txtCropMarkLength.Text = g_CropMarkLength
    txtStepRepeatPage.Text = g_StepRepeatPage
    txtStepRepeatSheetCount.Text = g_StepRepeatSheetCount

    RefreshStepRepeatVisibility

End Sub

Private Sub cboLayoutMode_Click()
    RefreshStepRepeatVisibility
End Sub

' Step & Repeat fields are only meaningful in that mode; grey them
' out otherwise so the form doesn't look like it needs values it won't use.
Private Sub RefreshStepRepeatVisibility()
    Dim isStepRepeat As Boolean
    isStepRepeat = (cboLayoutMode.Value = "StepRepeat")
    txtStepRepeatPage.Enabled = isStepRepeat
    txtStepRepeatSheetCount.Enabled = isStepRepeat
    lblStepRepeatPage.Enabled = isStepRepeat
    lblStepRepeatSheetCount.Enabled = isStepRepeat
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdRun_Click()

    If Not ValidateInputs() Then Exit Sub

    g_LayoutMode = cboLayoutMode.Value
    g_Unit = ComboIndexToUnit(cboUnit.ListIndex)

    g_SheetWidth = CDbl(txtSheetWidth.Text)
    g_SheetHeight = CDbl(txtSheetHeight.Text)
    g_GridRows = CInt(txtGridRows.Text)
    g_GridCols = CInt(txtGridCols.Text)
    g_GutterX = CDbl(txtGutterX.Text)
    g_GutterY = CDbl(txtGutterY.Text)
    g_MarginLeft = CDbl(txtMarginLeft.Text)
    g_MarginTop = CDbl(txtMarginTop.Text)
    g_BleedSize = CDbl(txtBleedSize.Text)
    g_AddCropMarks = (chkCropMarks.Value)
    g_CropMarkLength = CDbl(txtCropMarkLength.Text)
    g_StepRepeatPage = CInt(NzNum(txtStepRepeatPage.Text, 1))
    g_StepRepeatSheetCount = CInt(NzNum(txtStepRepeatSheetCount.Text, 1))

    Unload Me

    ' Persist only after a successful run. Guarded so a registry problem
    ' can never break the imposition itself. Cancel, failed validation
    ' (exited above) and a failed/erroring run never reach this save.
    If RunImposition() Then
        On Error Resume Next
        SaveCurrentSettings
        Err.Clear
        On Error GoTo 0
    End If

End Sub

'--------------------------------------------------------------
' Basic sanity checks before handing values off to RunImposition.
'--------------------------------------------------------------
Private Function ValidateInputs() As Boolean

    ValidateInputs = False

    If cboLayoutMode.ListIndex = -1 Then
        MsgBox "Please choose a Layout Mode.", vbExclamation
        Exit Function
    End If
    If cboUnit.ListIndex = -1 Then
        MsgBox "Please choose a Unit.", vbExclamation
        Exit Function
    End If
    If Not IsPositiveNumber(txtSheetWidth.Text) Or Not IsPositiveNumber(txtSheetHeight.Text) Then
        MsgBox "Sheet Width and Height must be positive numbers.", vbExclamation
        Exit Function
    End If
    If Not IsPositiveWholeNumber(txtGridRows.Text) Or Not IsPositiveWholeNumber(txtGridCols.Text) Then
        MsgBox "Grid Rows and Cols must be whole numbers of 1 or more.", vbExclamation
        Exit Function
    End If
    If Not IsNonNegativeNumber(txtGutterX.Text) Or Not IsNonNegativeNumber(txtGutterY.Text) Then
        MsgBox "Gutter values must be 0 or more.", vbExclamation
        Exit Function
    End If
    If Not IsNonNegativeNumber(txtMarginLeft.Text) Or Not IsNonNegativeNumber(txtMarginTop.Text) Then
        MsgBox "Margin values must be 0 or more.", vbExclamation
        Exit Function
    End If
    If Not IsNonNegativeNumber(txtBleedSize.Text) Then
        MsgBox "Bleed size must be 0 or more.", vbExclamation
        Exit Function
    End If
    If Not IsNonNegativeNumber(txtCropMarkLength.Text) Then
        MsgBox "Crop mark length must be 0 or more.", vbExclamation
        Exit Function
    End If
    If cboLayoutMode.Value = "StepRepeat" Then
        If Not IsPositiveWholeNumber(txtStepRepeatPage.Text) Then
            MsgBox "Step & Repeat page number must be a whole number of 1 or more.", vbExclamation
            Exit Function
        End If
        If Not IsPositiveWholeNumber(txtStepRepeatSheetCount.Text) Then
            MsgBox "Sheets to produce must be a whole number of 1 or more.", vbExclamation
            Exit Function
        End If
    End If
    If cboLayoutMode.Value = "Signature" Then
        If CInt(txtGridRows.Text) <> 1 Or CInt(txtGridCols.Text) <> 2 Then
            MsgBox "Signature mode currently requires Grid Rows = 1 and Grid Cols = 2.", vbExclamation
            Exit Function
        End If
    End If

    ValidateInputs = True

End Function

Private Function IsPositiveNumber(s As String) As Boolean
    IsPositiveNumber = IsNumeric(s) And CDbl(s) > 0
End Function

Private Function IsNonNegativeNumber(s As String) As Boolean
    IsNonNegativeNumber = IsNumeric(s) And CDbl(s) >= 0
End Function

Private Function IsPositiveWholeNumber(s As String) As Boolean
    IsPositiveWholeNumber = IsNumeric(s) And CDbl(s) >= 1 And (CDbl(s) = Int(CDbl(s)))
End Function

Private Function NzNum(s As String, defaultVal As Double) As Double
    If IsNumeric(s) Then
        NzNum = CDbl(s)
    Else
        NzNum = defaultVal
    End If
End Function

Private Function IndexOfString(cbo As Object, s As String) As Integer
    Dim i As Integer
    For i = 0 To cbo.ListCount - 1
        If cbo.List(i) = s Then
            IndexOfString = i
            Exit Function
        End If
    Next i
    IndexOfString = 0
End Function

Private Function UnitToComboIndex(u As cdrUnit) As Integer
    Select Case u
        Case cdrMillimeter: UnitToComboIndex = 0
        Case cdrInch:       UnitToComboIndex = 1
        Case cdrPoint:      UnitToComboIndex = 2
        Case cdrPixel:      UnitToComboIndex = 3
        Case Else:          UnitToComboIndex = 0
    End Select
End Function

Private Function ComboIndexToUnit(idx As Integer) As cdrUnit
    Select Case idx
        Case 0: ComboIndexToUnit = cdrMillimeter
        Case 1: ComboIndexToUnit = cdrInch
        Case 2: ComboIndexToUnit = cdrPoint
        Case 3: ComboIndexToUnit = cdrPixel
        Case Else: ComboIndexToUnit = cdrMillimeter
    End Select
End Function

