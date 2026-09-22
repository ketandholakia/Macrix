VERSION 5.00
Begin VB.Form frmImpositionSettings 
   Caption         =   "Imposition Settings"
   ClientHeight    =   6180
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   6180
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2520
      TabIndex        =   17
      Top             =   5680
      Width           =   1215
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Run Imposition"
      Default         =   -1  'True
      Height          =   375
      Left            =   960
      TabIndex        =   16
      Top             =   5680
      Width           =   1455
   End
   Begin VB.TextBox txtStepRepeatSheetCount 
      Height          =   285
      Left            =   2520
      TabIndex        =   15
      Top             =   5240
      Width           =   1215
   End
   Begin VB.Label lblStepRepeatSheetCount 
      Caption         =   "Sheets to produce"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   5280
      Width           =   2175
   End
   Begin VB.TextBox txtStepRepeatPage 
      Height          =   285
      Left            =   2520
      TabIndex        =   13
      Top             =   4880
      Width           =   1215
   End
   Begin VB.Label lblStepRepeatPage 
      Caption         =   "Repeat which page #"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   4920
      Width           =   2175
   End
   Begin VB.CheckBox chkCropMarks 
      Caption         =   "Add crop marks"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   4480
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.TextBox txtCropMarkLength 
      Height          =   285
      Left            =   2520
      TabIndex        =   10
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox txtBleedSize 
      Height          =   285
      Left            =   2520
      TabIndex        =   9
      Top             =   4040
      Width           =   1215
   End
   Begin VB.Label lblBleedSize 
      Caption         =   "Bleed size"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   4080
      Width           =   2175
   End
   Begin VB.TextBox txtMarginTop 
      Height          =   285
      Left            =   2520
      TabIndex        =   7
      Top             =   3640
      Width           =   1215
   End
   Begin VB.TextBox txtMarginLeft 
      Height          =   285
      Left            =   960
      TabIndex        =   6
      Top             =   3640
      Width           =   1215
   End
   Begin VB.Label lblMargins 
      Caption         =   "Margin L / T"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   3320
      Width           =   2175
   End
   Begin VB.TextBox txtGutterY 
      Height          =   285
      Left            =   2520
      TabIndex        =   4
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox txtGutterX 
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label lblGutters 
      Caption         =   "Gutter X / Y"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   2680
      Width           =   2175
   End
   Begin VB.TextBox txtGridCols 
      Height          =   285
      Left            =   2520
      TabIndex        =   1
      Top             =   2360
      Width           =   1215
   End
   Begin VB.TextBox txtGridRows 
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   2360
      Width           =   1215
   End
   Begin VB.Label lblGrid 
      Caption         =   "Grid Rows / Cols"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   2040
      Width           =   2175
   End
   Begin VB.TextBox txtSheetHeight 
      Height          =   285
      Left            =   2520
      TabIndex        =   19
      Top             =   1720
      Width           =   1215
   End
   Begin VB.TextBox txtSheetWidth 
      Height          =   285
      Left            =   960
      TabIndex        =   20
      Top             =   1720
      Width           =   1215
   End
   Begin VB.Label lblSheetSize 
      Caption         =   "Sheet Width / Height"
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   1400
      Width           =   2295
   End
   Begin VB.ComboBox cboUnit 
      Height          =   315
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   1000
      Width           =   1815
   End
   Begin VB.Label lblUnit 
      Caption         =   "Unit"
      Height          =   255
      Left            =   240
      TabIndex        =   23
      Top             =   1040
      Width           =   2175
   End
   Begin VB.ComboBox cboLayoutMode 
      Height          =   315
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   560
      Width           =   1815
   End
   Begin VB.Label lblLayoutMode 
      Caption         =   "Layout Mode"
      Height          =   255
      Left            =   240
      TabIndex        =   25
      Top             =   600
      Width           =   2175
   End
End
Attribute VB_Name = "frmImpositionSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================
' frmImpositionSettings
' Simple settings UI over modSettings' g_ variables. On "Run
' Imposition" it validates input, writes it back to the g_
' variables, and calls modImposition.RunImposition.
'==============================================================

Option Explicit

Private Sub UserForm_Initialize()
    ' Not used for VB.Form (CorelDRAW VBA); Form_Load below does the work.
End Sub

Private Sub Form_Load()

    If Not g_SettingsInitialized Then InitDefaultSettings

    cboLayoutMode.AddItem "Simple"
    cboLayoutMode.AddItem "StepRepeat"
    cboLayoutMode.AddItem "Signature"
    cboLayoutMode.ListIndex = IndexOfString(cboLayoutMode, g_LayoutMode)

    ' NOTE: enum member names (cdrMillimeter, cdrInch, etc.) can differ
    ' slightly by CorelDRAW version. If any of these four don't compile,
    ' check the exact names via the VBA Object Browser (F2, search "cdrUnit").
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
    isStepRepeat = (cboLayoutMode.Text = "StepRepeat")
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

    g_LayoutMode = cboLayoutMode.Text
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
    g_AddCropMarks = (chkCropMarks.Value = 1)
    g_CropMarkLength = CDbl(txtCropMarkLength.Text)
    g_StepRepeatPage = CInt(NzNum(txtStepRepeatPage.Text, 1))
    g_StepRepeatSheetCount = CInt(NzNum(txtStepRepeatSheetCount.Text, 1))

    Unload Me
    RunImposition

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
    If cboLayoutMode.Text = "StepRepeat" Then
        If Not IsPositiveWholeNumber(txtStepRepeatPage.Text) Then
            MsgBox "Step & Repeat page number must be a whole number of 1 or more.", vbExclamation
            Exit Function
        End If
        If Not IsPositiveWholeNumber(txtStepRepeatSheetCount.Text) Then
            MsgBox "Sheets to produce must be a whole number of 1 or more.", vbExclamation
            Exit Function
        End If
    End If
    If cboLayoutMode.Text = "Signature" Then
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

Private Function IndexOfString(cbo As ComboBox, s As String) As Integer
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
