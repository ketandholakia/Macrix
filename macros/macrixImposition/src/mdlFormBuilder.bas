Attribute VB_Name = "mdlFormBuilder"
'==============================================================
' mdlFormBuilder (macrixImposition)
'
' Builds frmImpositionSettings as a real MSForms VBA UserForm at
' runtime via the VBE Extensibility model. This replaces importing
' the legacy VB6-format .frm file, which CorelDRAW's VBE cannot
' load ("The form class contained in the file is not supported").
'
' The injected code-behind mirrors forms/frmImpositionSettings.frm,
' with one required change: VB6's Form_Load is replaced by
' UserForm_Initialize (VBA UserForms never fire Form_Load).
'
' Coordinate conversion: VB6 twips / 15 = MSForms points.
'==============================================================
Option Explicit
Option Private Module

Private Const FORM_NAME As String = "frmImpositionSettings"

Public Sub BuildImpositionFormWithCode()

    Dim vbProj As Object
    Set vbProj = TargetProject()
    If vbProj Is Nothing Then
        MsgBox "Could not access the VBA project model." & vbCrLf & vbCrLf & _
               "CorelDRAW must allow VBA project access (Tools > Options > VBA).", _
               vbCritical, "Imposition"
        Exit Sub
    End If

    ' 1. Remove any existing/broken form to prevent duplicates
    On Error Resume Next
    Dim existingComp As Object
    Set existingComp = vbProj.VBComponents(FORM_NAME)
    If Not existingComp Is Nothing Then
        vbProj.VBComponents.Remove existingComp
    End If
    Err.Clear
    On Error GoTo 0

    ' 2. Add the UserForm component (3 = MSForm)
    Dim newFormComp As Object
    Set newFormComp = vbProj.VBComponents.Add(3)
    ' Guarded: VBComponents.Add can return Nothing, and the next line would then raise
    ' "Run-time error 424: Object required" with no clue as to why.
    If newFormComp Is Nothing Then
        MsgBox "Could not create the UserForm component (project may be read-only or locked).", _
               vbCritical, "Imposition"
        Exit Sub
    End If
    newFormComp.Name = FORM_NAME

    ' 3. Configure the form itself
    With newFormComp.Properties
        .Item("Caption").Value = "Imposition Settings"
        .Item("Width").Value = 310
        .Item("Height").Value = 452
    End With

    Dim d As Object
    Set d = newFormComp.Designer

    ' 4. Add UI controls (positions converted from the .frm, twips/15)
    AddLabel d, "lblLayoutMode", "Layout Mode", 16, 40, 145, 17
    AddComboList d, "cboLayoutMode", 168, 37, 121, 20

    AddLabel d, "lblUnit", "Unit", 16, 69, 145, 17
    AddComboList d, "cboUnit", 168, 67, 121, 20

    AddLabel d, "lblSheetSize", "Sheet Width / Height", 16, 93, 153, 17
    AddTextBox d, "txtSheetWidth", "", 64, 115, 81, 19
    AddTextBox d, "txtSheetHeight", "", 168, 115, 81, 19

    AddLabel d, "lblGrid", "Grid Rows / Cols", 16, 136, 145, 17
    AddTextBox d, "txtGridRows", "", 64, 157, 81, 19
    AddTextBox d, "txtGridCols", "", 168, 157, 81, 19

    AddLabel d, "lblGutters", "Gutter X / Y", 16, 179, 145, 17
    AddTextBox d, "txtGutterX", "", 64, 200, 81, 19
    AddTextBox d, "txtGutterY", "", 168, 200, 81, 19

    AddLabel d, "lblMargins", "Margin L / T", 16, 221, 145, 17
    AddTextBox d, "txtMarginLeft", "", 64, 243, 81, 19
    AddTextBox d, "txtMarginTop", "", 168, 243, 81, 19

    AddLabel d, "lblBleedSize", "Bleed size", 16, 272, 145, 17
    AddTextBox d, "txtBleedSize", "", 168, 269, 81, 19

    AddCheckBox d, "chkCropMarks", "Add crop marks", 16, 299, 145, True
    AddTextBox d, "txtCropMarkLength", "", 168, 296, 81, 19

    AddLabel d, "lblStepRepeatPage", "Repeat which page #", 16, 328, 145, 17
    AddTextBox d, "txtStepRepeatPage", "", 168, 325, 81, 19

    AddLabel d, "lblStepRepeatSheetCount", "Sheets to produce", 16, 352, 145, 17
    AddTextBox d, "txtStepRepeatSheetCount", "", 168, 349, 81, 19

    AddButton d, "cmdRun", "Run Imposition", 64, 379, 97, 25
    AddButton d, "cmdCancel", "Cancel", 168, 379, 81, 25

    ' 5. Inject code-behind
    InjectCodeBehind newFormComp.CodeModule

End Sub

' Returns True if a form component with the expected name exists.
Public Function ImpositionFormExists() As Boolean
    On Error Resume Next
    Dim proj As Object
    Set proj = TargetProject()
    If proj Is Nothing Then Exit Function
    ImpositionFormExists = Not proj.VBComponents(FORM_NAME) Is Nothing
    Err.Clear
    On Error GoTo 0
End Function

' Resolve the VBA project this code lives in. Application.VBE.ActiveVBProject is
' frequently Nothing at runtime (the editor is not open), which silently broke the
' form build -- so find the project that owns this module instead.
Private Function TargetProject() As Object
    On Error Resume Next
    Dim proj As Object, comp As Object
    Set TargetProject = Application.VBE.ActiveVBProject
    If Not TargetProject Is Nothing Then Exit Function
    Set TargetProject = Nothing
    For Each proj In Application.VBE.VBProjects
        For Each comp In proj.VBComponents
            If StrComp(comp.Name, "mdlFormBuilder", vbTextCompare) = 0 Then
                Set TargetProject = proj
                Exit Function
            End If
        Next comp
    Next proj
    Err.Clear
    On Error GoTo 0
End Function

'--------------------------------------------------------------
' Control helpers
'--------------------------------------------------------------
Private Function AddComboList(ByVal container As Object, ByVal cName As String, _
                              ByVal cLeft As Single, ByVal cTop As Single, _
                              ByVal cWidth As Single, ByVal cHeight As Single) As Object
    Set AddComboList = container.Controls.Add("Forms.ComboBox.1", cName)
    With AddComboList
        .Left = cLeft: .Top = cTop: .Width = cWidth: .Height = cHeight
        .Style = 2 ' fmStyleDropDownList - matches VB6 Style=2 (Dropdown List)
    End With
End Function

Private Function AddCheckBox(ByVal container As Object, ByVal cName As String, ByVal cCaption As String, _
                             ByVal cLeft As Single, ByVal cTop As Single, ByVal cWidth As Single, _
                             ByVal cVal As Boolean) As Object
    Set AddCheckBox = container.Controls.Add("Forms.CheckBox.1", cName)
    With AddCheckBox: .Caption = cCaption: .Left = cLeft: .Top = cTop: .Width = cWidth: .Height = 17: .Value = cVal: End With
End Function

Private Function AddButton(ByVal container As Object, ByVal cName As String, ByVal cCaption As String, _
                           ByVal cLeft As Single, ByVal cTop As Single, ByVal cWidth As Single, _
                           ByVal cHeight As Single) As Object
    Set AddButton = container.Controls.Add("Forms.CommandButton.1", cName)
    With AddButton: .Caption = cCaption: .Left = cLeft: .Top = cTop: .Width = cWidth: .Height = cHeight: End With
End Function

Private Function AddTextBox(ByVal container As Object, ByVal cName As String, ByVal cText As String, _
                            ByVal cLeft As Single, ByVal cTop As Single, ByVal cWidth As Single, _
                            ByVal cHeight As Single) As Object
    Set AddTextBox = container.Controls.Add("Forms.TextBox.1", cName)
    With AddTextBox: .Text = cText: .Left = cLeft: .Top = cTop: .Width = cWidth: .Height = cHeight: End With
End Function

Private Function AddLabel(ByVal container As Object, ByVal cName As String, ByVal cCaption As String, _
                          ByVal cLeft As Single, ByVal cTop As Single, ByVal cWidth As Single, _
                          ByVal cHeight As Single) As Object
    Set AddLabel = container.Controls.Add("Forms.Label.1", cName)
    With AddLabel: .Caption = cCaption: .Left = cLeft: .Top = cTop: .Width = cWidth: .Height = cHeight: End With
End Function

'--------------------------------------------------------------
' Code-behind injection.
' NOTE: built line-by-line with s = s & ... to stay far below the
' VBA limit of 24 line continuations per statement (see
' docs/VBA-Form-AutoBuilder.md).
'--------------------------------------------------------------
Private Sub InjectCodeBehind(ByVal codeMod As Object)
    Dim s As String
    s = "Option Explicit" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_Initialize()" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If Not g_SettingsInitialized Then InitDefaultSettings" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    cboLayoutMode.AddItem ""Simple""" & vbCrLf
    s = s & "    cboLayoutMode.AddItem ""StepRepeat""" & vbCrLf
    s = s & "    cboLayoutMode.AddItem ""Signature""" & vbCrLf
    s = s & "    cboLayoutMode.ListIndex = IndexOfString(cboLayoutMode, g_LayoutMode)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    cboUnit.AddItem ""Millimeters""" & vbCrLf
    s = s & "    cboUnit.AddItem ""Inches""" & vbCrLf
    s = s & "    cboUnit.AddItem ""Points""" & vbCrLf
    s = s & "    cboUnit.AddItem ""Pixels""" & vbCrLf
    s = s & "    cboUnit.ListIndex = UnitToComboIndex(g_Unit)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    txtSheetWidth.Text = g_SheetWidth" & vbCrLf
    s = s & "    txtSheetHeight.Text = g_SheetHeight" & vbCrLf
    s = s & "    txtGridRows.Text = g_GridRows" & vbCrLf
    s = s & "    txtGridCols.Text = g_GridCols" & vbCrLf
    s = s & "    txtGutterX.Text = g_GutterX" & vbCrLf
    s = s & "    txtGutterY.Text = g_GutterY" & vbCrLf
    s = s & "    txtMarginLeft.Text = g_MarginLeft" & vbCrLf
    s = s & "    txtMarginTop.Text = g_MarginTop" & vbCrLf
    s = s & "    txtBleedSize.Text = g_BleedSize" & vbCrLf
    s = s & "    chkCropMarks.Value = IIf(g_AddCropMarks, 1, 0)" & vbCrLf
    s = s & "    txtCropMarkLength.Text = g_CropMarkLength" & vbCrLf
    s = s & "    txtStepRepeatPage.Text = g_StepRepeatPage" & vbCrLf
    s = s & "    txtStepRepeatSheetCount.Text = g_StepRepeatSheetCount" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    RefreshStepRepeatVisibility" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cboLayoutMode_Click()" & vbCrLf
    s = s & "    RefreshStepRepeatVisibility" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' Step & Repeat fields are only meaningful in that mode; grey them" & vbCrLf
    s = s & "' out otherwise so the form doesn't look like it needs values it won't use." & vbCrLf
    s = s & "Private Sub RefreshStepRepeatVisibility()" & vbCrLf
    s = s & "    Dim isStepRepeat As Boolean" & vbCrLf
    s = s & "    isStepRepeat = (cboLayoutMode.Value = ""StepRepeat"")" & vbCrLf
    s = s & "    txtStepRepeatPage.Enabled = isStepRepeat" & vbCrLf
    s = s & "    txtStepRepeatSheetCount.Enabled = isStepRepeat" & vbCrLf
    s = s & "    lblStepRepeatPage.Enabled = isStepRepeat" & vbCrLf
    s = s & "    lblStepRepeatSheetCount.Enabled = isStepRepeat" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdCancel_Click()" & vbCrLf
    s = s & "    Unload Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdRun_Click()" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If Not ValidateInputs() Then Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    g_LayoutMode = cboLayoutMode.Value" & vbCrLf
    s = s & "    g_Unit = ComboIndexToUnit(cboUnit.ListIndex)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    g_SheetWidth = CDbl(txtSheetWidth.Text)" & vbCrLf
    s = s & "    g_SheetHeight = CDbl(txtSheetHeight.Text)" & vbCrLf
    s = s & "    g_GridRows = CInt(txtGridRows.Text)" & vbCrLf
    s = s & "    g_GridCols = CInt(txtGridCols.Text)" & vbCrLf
    s = s & "    g_GutterX = CDbl(txtGutterX.Text)" & vbCrLf
    s = s & "    g_GutterY = CDbl(txtGutterY.Text)" & vbCrLf
    s = s & "    g_MarginLeft = CDbl(txtMarginLeft.Text)" & vbCrLf
    s = s & "    g_MarginTop = CDbl(txtMarginTop.Text)" & vbCrLf
    s = s & "    g_BleedSize = CDbl(txtBleedSize.Text)" & vbCrLf
    s = s & "    g_AddCropMarks = (chkCropMarks.Value)" & vbCrLf
    s = s & "    g_CropMarkLength = CDbl(txtCropMarkLength.Text)" & vbCrLf
    s = s & "    g_StepRepeatPage = CInt(NzNum(txtStepRepeatPage.Text, 1))" & vbCrLf
    s = s & "    g_StepRepeatSheetCount = CInt(NzNum(txtStepRepeatSheetCount.Text, 1))" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Unload Me" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ' Persist only after a successful run. Guarded so a registry problem" & vbCrLf
    s = s & "    ' can never break the imposition itself. Cancel, failed validation" & vbCrLf
    s = s & "    ' (exited above) and a failed/erroring run never reach this save." & vbCrLf
    s = s & "    If RunImposition() Then" & vbCrLf
    s = s & "        On Error Resume Next" & vbCrLf
    s = s & "        SaveCurrentSettings" & vbCrLf
    s = s & "        Err.Clear" & vbCrLf
    s = s & "        On Error GoTo 0" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'--------------------------------------------------------------" & vbCrLf
    s = s & "' Basic sanity checks before handing values off to RunImposition." & vbCrLf
    s = s & "'--------------------------------------------------------------" & vbCrLf
    s = s & "Private Function ValidateInputs() As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ValidateInputs = False" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If cboLayoutMode.ListIndex = -1 Then" & vbCrLf
    s = s & "        MsgBox ""Please choose a Layout Mode."", vbExclamation" & vbCrLf
    s = s & "        Exit Function" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If cboUnit.ListIndex = -1 Then" & vbCrLf
    s = s & "        MsgBox ""Please choose a Unit."", vbExclamation" & vbCrLf
    s = s & "        Exit Function" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If Not IsPositiveNumber(txtSheetWidth.Text) Or Not IsPositiveNumber(txtSheetHeight.Text) Then" & vbCrLf
    s = s & "        MsgBox ""Sheet Width and Height must be positive numbers."", vbExclamation" & vbCrLf
    s = s & "        Exit Function" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If Not IsPositiveWholeNumber(txtGridRows.Text) Or Not IsPositiveWholeNumber(txtGridCols.Text) Then" & vbCrLf
    s = s & "        MsgBox ""Grid Rows and Cols must be whole numbers of 1 or more."", vbExclamation" & vbCrLf
    s = s & "        Exit Function" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If Not IsNonNegativeNumber(txtGutterX.Text) Or Not IsNonNegativeNumber(txtGutterY.Text) Then" & vbCrLf
    s = s & "        MsgBox ""Gutter values must be 0 or more."", vbExclamation" & vbCrLf
    s = s & "        Exit Function" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If Not IsNonNegativeNumber(txtMarginLeft.Text) Or Not IsNonNegativeNumber(txtMarginTop.Text) Then" & vbCrLf
    s = s & "        MsgBox ""Margin values must be 0 or more."", vbExclamation" & vbCrLf
    s = s & "        Exit Function" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If Not IsNonNegativeNumber(txtBleedSize.Text) Then" & vbCrLf
    s = s & "        MsgBox ""Bleed size must be 0 or more."", vbExclamation" & vbCrLf
    s = s & "        Exit Function" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If Not IsNonNegativeNumber(txtCropMarkLength.Text) Then" & vbCrLf
    s = s & "        MsgBox ""Crop mark length must be 0 or more."", vbExclamation" & vbCrLf
    s = s & "        Exit Function" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If cboLayoutMode.Value = ""StepRepeat"" Then" & vbCrLf
    s = s & "        If Not IsPositiveWholeNumber(txtStepRepeatPage.Text) Then" & vbCrLf
    s = s & "            MsgBox ""Step & Repeat page number must be a whole number of 1 or more."", vbExclamation" & vbCrLf
    s = s & "            Exit Function" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "        If Not IsPositiveWholeNumber(txtStepRepeatSheetCount.Text) Then" & vbCrLf
    s = s & "            MsgBox ""Sheets to produce must be a whole number of 1 or more."", vbExclamation" & vbCrLf
    s = s & "            Exit Function" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If cboLayoutMode.Value = ""Signature"" Then" & vbCrLf
    s = s & "        If CInt(txtGridRows.Text) <> 1 Or CInt(txtGridCols.Text) <> 2 Then" & vbCrLf
    s = s & "            MsgBox ""Signature mode currently requires Grid Rows = 1 and Grid Cols = 2."", vbExclamation" & vbCrLf
    s = s & "            Exit Function" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ValidateInputs = True" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function IsPositiveNumber(s As String) As Boolean" & vbCrLf
    s = s & "    IsPositiveNumber = IsNumeric(s) And CDbl(s) > 0" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function IsNonNegativeNumber(s As String) As Boolean" & vbCrLf
    s = s & "    IsNonNegativeNumber = IsNumeric(s) And CDbl(s) >= 0" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function IsPositiveWholeNumber(s As String) As Boolean" & vbCrLf
    s = s & "    IsPositiveWholeNumber = IsNumeric(s) And CDbl(s) >= 1 And (CDbl(s) = Int(CDbl(s)))" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function NzNum(s As String, defaultVal As Double) As Double" & vbCrLf
    s = s & "    If IsNumeric(s) Then" & vbCrLf
    s = s & "        NzNum = CDbl(s)" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        NzNum = defaultVal" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function IndexOfString(cbo As Object, s As String) As Integer" & vbCrLf
    s = s & "    Dim i As Integer" & vbCrLf
    s = s & "    For i = 0 To cbo.ListCount - 1" & vbCrLf
    s = s & "        If cbo.List(i) = s Then" & vbCrLf
    s = s & "            IndexOfString = i" & vbCrLf
    s = s & "            Exit Function" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "    IndexOfString = 0" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function UnitToComboIndex(u As cdrUnit) As Integer" & vbCrLf
    s = s & "    Select Case u" & vbCrLf
    s = s & "        Case cdrMillimeter: UnitToComboIndex = 0" & vbCrLf
    s = s & "        Case cdrInch:       UnitToComboIndex = 1" & vbCrLf
    s = s & "        Case cdrPoint:      UnitToComboIndex = 2" & vbCrLf
    s = s & "        Case cdrPixel:      UnitToComboIndex = 3" & vbCrLf
    s = s & "        Case Else:          UnitToComboIndex = 0" & vbCrLf
    s = s & "    End Select" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function ComboIndexToUnit(idx As Integer) As cdrUnit" & vbCrLf
    s = s & "    Select Case idx" & vbCrLf
    s = s & "        Case 0: ComboIndexToUnit = cdrMillimeter" & vbCrLf
    s = s & "        Case 1: ComboIndexToUnit = cdrInch" & vbCrLf
    s = s & "        Case 2: ComboIndexToUnit = cdrPoint" & vbCrLf
    s = s & "        Case 3: ComboIndexToUnit = cdrPixel" & vbCrLf
    s = s & "        Case Else: ComboIndexToUnit = cdrMillimeter" & vbCrLf
    s = s & "    End Select" & vbCrLf
    s = s & "End Function" & vbCrLf

    codeMod.DeleteLines 1, codeMod.CountOfLines
    codeMod.AddFromString s
End Sub
