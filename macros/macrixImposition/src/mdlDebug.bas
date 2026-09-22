Attribute VB_Name = "mdlDebugLog"
Option Explicit

Sub DebugBuildForm()
    Dim f As Integer
    f = FreeFile
    Open "C:\Users\Admin\AppData\Local\Temp\vba-debug.log" For Output As #f
    Print #f, "start"
    On Error Resume Next
    Dim p As Object
    Set p = Application.VBE.ActiveVBProject
    Print #f, "ActiveVBProject err=" & Err.Number & " name=" & p.Name
    Err.Clear
    Dim vbp As Object
    For Each vbp In Application.VBE.VBProjects
        If vbp.Name = "macrixImposition" Then Set p = vbp
    Next vbp
    Print #f, "target found: " & (Not p Is Nothing)
    Err.Clear
    Dim comp As Object
    Set comp = p.VBComponents("mdlFormBuilder")
    Print #f, "mdlFormBuilder present: " & (Not comp Is Nothing) & " err=" & Err.Number
    Err.Clear
    Dim newComp As Object
    Set newComp = p.VBComponents.Add(3)
    Print #f, "Add(3) err=" & Err.Number & " name=" & newComp.Name
    Err.Clear
    newComp.Name = "frmDebugTest"
    Print #f, "rename err=" & Err.Number
    Err.Clear
    newComp.Properties.Item("Caption").Value = "test"
    Print #f, "caption err=" & Err.Number
    Err.Clear
    Dim d As Object
    Set d = newComp.Designer
    Print #f, "designer err=" & Err.Number
    Err.Clear
    Dim ctl As Object
    Set ctl = d.Controls.Add("Forms.TextBox.1", "txtTest")
    Print #f, "control add err=" & Err.Number
    Err.Clear
    p.VBComponents.Remove newComp
    Print #f, "cleanup err=" & Err.Number
    Err.Clear
    On Error GoTo 0
    Print #f, "calling builder"
    On Error Resume Next
    mdlFormBuilder.BuildImpositionFormWithCode
    Print #f, "builder err=" & Err.Number & " desc=" & Err.Description
    On Error GoTo 0
    Print #f, "done"
    Close #f
End Sub
