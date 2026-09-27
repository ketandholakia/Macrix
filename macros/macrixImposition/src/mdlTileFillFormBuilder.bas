Attribute VB_Name = "mdlTileFillFormBuilder"
'==============================================================
' mdlTileFillFormBuilder (macrixImposition project)
'
' Builds frmTileFillSettings as a real MSForms VBA UserForm at
' runtime via the VBE Extensibility model -- same reason and same
' approach as mdlFormBuilder/frmImpositionSettings: CorelDRAW's VBE
' cannot load a legacy VB6-format .frm file directly.
'
' Coordinate convention matches mdlFormBuilder: values below are
' already in MSForms points (not VB6 twips), so they can be used
' directly.
'==============================================================

Option Explicit
Option Private Module

Private Const FORM_NAME As String = "frmTileFillSettings"

Public Sub BuildTileFillFormWithCode()

    Dim vbProj As Object
    Set vbProj = TargetProject()
    If vbProj Is Nothing Then
        MsgBox "Could not access the VBA project model." & vbCrLf & vbCrLf & _
               "CorelDRAW must allow VBA project access (Tools > Options > VBA).", _
               vbCritical, "Tile Fill"
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
    If newFormComp Is Nothing Then
        MsgBox "Could not create the UserForm component (project may be read-only or locked).", _
               vbCritical, "Tile Fill"
        Exit Sub
    End If
    newFormComp.Name = FORM_NAME

    ' 3. Configure the form itself
    With newFormComp.Properties
        .Item("Caption").Value = "Tile Fill Settings"
        .Item("Width").Value = 300
        .Item("Height").Value = 270
    End With

    Dim d As Object
    Set d = newFormComp.Designer

    ' 4. Add UI controls
    AddLabel d, "lblUnit", "Unit", 16, 16, 145, 17
    AddComboList d, "cboUnit", 168, 13, 105, 20

    AddLabel d, "lblMargin", "Margin X / Y", 16, 45, 153, 17
    AddTextBox d, "txtMarginX", "", 64, 67, 81, 19
    AddTextBox d, "txtMarginY", "", 168, 67, 81, 19

    AddLabel d, "lblGutter", "Gutter X / Y", 16, 92, 153, 17
    AddTextBox d, "txtGutterX", "", 64, 114, 81, 19
    AddTextBox d, "txtGutterY", "", 168, 114, 81, 19

    AddCheckBox d, "chkAllowRotate", "Allow 90-degree rotation to fit more", 16, 145, 240, True

    AddButton d, "cmdRun", "Tile Fill Selection", 40, 178, 130, 25
    AddButton d, "cmdCancel", "Cancel", 176, 178, 81, 25

    ' 5. Inject code-behind
    InjectCodeBehind newFormComp.CodeModule

End Sub

' Returns True if a form component with the expected name exists.
Public Function TileFillFormExists() As Boolean
    On Error Resume Next
    Dim proj As Object
    Set proj = TargetProject()
    If proj Is Nothing Then Exit Function
    TileFillFormExists = Not proj.VBComponents(FORM_NAME) Is Nothing
    Err.Clear
    On Error GoTo 0
End Function

' Resolve the VBA project this code lives in. Application.VBE.ActiveVBProject is
' frequently Nothing at runtime (the editor is not open), so find the project
' that owns this module instead -- same fix as mdlFormBuilder's TargetProject.
Private Function TargetProject() As Object
    On Error Resume Next
    Dim proj As Object, comp As Object
    Set TargetProject = Application.VBE.ActiveVBProject
    If Not TargetProject Is Nothing Then Exit Function
    Set TargetProject = Nothing
    For Each proj In Application.VBE.VBProjects
        For Each comp In proj.VBComponents
            If StrComp(comp.Name, "mdlTileFillFormBuilder", vbTextCompare) = 0 Then
                Set TargetProject = proj
                Exit Function
            End If
        Next comp
    Next proj
    Err.Clear
    On Error GoTo 0
End Function

'--------------------------------------------------------------
' Control helpers (Private -- same names as mdlFormBuilder's own
' helpers are fine, module-private scope keeps them from colliding)
'--------------------------------------------------------------
Private Function AddComboList(ByVal container As Object, ByVal cName As String, _
                              ByVal cLeft As Single, ByVal cTop As Single, _
                              ByVal cWidth As Single, ByVal cHeight As Single) As Object
    Set AddComboList = container.Controls.Add("Forms.ComboBox.1", cName)
    With AddComboList
        .Left = cLeft: .Top = cTop: .Width = cWidth: .Height = cHeight
        .Style = 2 ' fmStyleDropDownList
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
' Code-behind injection. Built line-by-line (s = s & ...) to stay
' well below VBA's 24-line-continuation limit per statement, same
' as mdlFormBuilder's InjectCodeBehind.
'--------------------------------------------------------------
Private Sub InjectCodeBehind(ByVal codeMod As Object)
    Dim s As String
    s = "Option Explicit" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_Initialize()" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If Not g_TileSettingsInitialized Then InitDefaultTileSettings" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    cboUnit.AddItem ""Millimeters""" & vbCrLf
    s = s & "    cboUnit.AddItem ""Inches""" & vbCrLf
    s = s & "    cboUnit.AddItem ""Points""" & vbCrLf
    s = s & "    cboUnit.AddItem ""Pixels""" & vbCrLf
    s = s & "    cboUnit.ListIndex = UnitToComboIndex(g_TileUnit)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    txtMarginX.Text = g_TileMarginX" & vbCrLf
    s = s & "    txtMarginY.Text = g_TileMarginY" & vbCrLf
    s = s & "    txtGutterX.Text = g_TileGutterX" & vbCrLf
    s = s & "    txtGutterY.Text = g_TileGutterY" & vbCrLf
    s = s & "    chkAllowRotate.Value = IIf(g_TileAllowRotate, 1, 0)" & vbCrLf
    s = s & "" & vbCrLf
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
    s = s & "    g_TileUnit = ComboIndexToUnit(cboUnit.ListIndex)" & vbCrLf
    s = s & "    g_TileMarginX = CDbl(txtMarginX.Text)" & vbCrLf
    s = s & "    g_TileMarginY = CDbl(txtMarginY.Text)" & vbCrLf
    s = s & "    g_TileGutterX = CDbl(txtGutterX.Text)" & vbCrLf
    s = s & "    g_TileGutterY = CDbl(txtGutterY.Text)" & vbCrLf
    s = s & "    g_TileAllowRotate = (chkAllowRotate.Value)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Unload Me" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ' Persist only after a successful run -- same rule as" & vbCrLf
    s = s & "    ' frmImpositionSettings: Cancel, failed validation (exited" & vbCrLf
    s = s & "    ' above), and a failed/erroring run never reach this save." & vbCrLf
    s = s & "    If RunTileFill() Then" & vbCrLf
    s = s & "        On Error Resume Next" & vbCrLf
    s = s & "        SaveCurrentTileSettings" & vbCrLf
    s = s & "        Err.Clear" & vbCrLf
    s = s & "        On Error GoTo 0" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'--------------------------------------------------------------" & vbCrLf
    s = s & "' Basic sanity checks before handing values off to RunTileFill." & vbCrLf
    s = s & "'--------------------------------------------------------------" & vbCrLf
    s = s & "Private Function ValidateInputs() As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ValidateInputs = False" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If cboUnit.ListIndex = -1 Then" & vbCrLf
    s = s & "        MsgBox ""Please choose a Unit."", vbExclamation" & vbCrLf
    s = s & "        Exit Function" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If Not IsNonNegativeNumber(txtMarginX.Text) Or Not IsNonNegativeNumber(txtMarginY.Text) Then" & vbCrLf
    s = s & "        MsgBox ""Margin values must be 0 or more."", vbExclamation" & vbCrLf
    s = s & "        Exit Function" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If Not IsNonNegativeNumber(txtGutterX.Text) Or Not IsNonNegativeNumber(txtGutterY.Text) Then" & vbCrLf
    s = s & "        MsgBox ""Gutter values must be 0 or more."", vbExclamation" & vbCrLf
    s = s & "        Exit Function" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ValidateInputs = True" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function IsNonNegativeNumber(s As String) As Boolean" & vbCrLf
    s = s & "    IsNonNegativeNumber = IsNumeric(s) And CDbl(s) >= 0" & vbCrLf
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
