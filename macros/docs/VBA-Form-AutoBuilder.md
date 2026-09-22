# VBA Form Auto-Builder Guide

> **Status note:** `vittixBleed` originally shipped a `mdlFormBuilder.bas` using this
> approach, but it was never wired up (nothing called `BuildMainViewFormWithCode`) and
> it was incomplete - it didn't create the `cbRound`/`obRound0-2` controls that
> `MainLogic.Create` expects, and its injected button captions were placeholder text
> rather than real strings. `Sync-To-GMS.ps1` imports `MainView.frm` directly and that
> works fine, so the file was removed rather than finished. If the `.frm` import error
> described below ever resurfaces for `vittixBleed` or a future macro, revive this
> pattern from the boilerplate below - just make sure the generated control set exactly
> matches what the form's code-behind expects, and use real caption text, not
> placeholders.

When migrating macros from VB6 (or modern twinBASIC environments) to CorelDRAW's VBA Editor (VBE), you often encounter the error:
> **"The form class contained in the file is not supported in VBE. The file can't be loaded."**

This happens because VBA UserForms are proprietary `MSForms.UserForm` binaries that store layout in `.frx` blobs, whereas VB6 forms (`VB.Form`) store layout as plain text in the `.frm` file.

To bypass hours of manual drag-and-drop recreation in CorelDRAW, we use **VBA Extensibility** to programmatically generate the form UI and inject the code-behind.

## The Strategy

Instead of trying to import the `.frm` file directly, we create a standard module (`mdlFormBuilder.bas`) that uses the `VBProject.VBComponents.Add(3)` method to build the form on the fly.

### Important Rules to Remember

1. **Reference Requirement:** Any macro project using the builder script must have **Tools → References → Microsoft Visual Basic for Applications Extensibility 5.3** checked.
2. **Line Continuation Limit (CRITICAL):** When injecting the code-behind as a string, **do not** use endless line continuations (`& _`). VBA has a hard-coded compiler limit of **24 line continuations** per statement. If you exceed this, the VBA editor will secretly treat the `.bas` file as having a fatal syntax error and throw `Exception from HRESULT: 0x800A9D00` during import. 
   - **Bad:** `s = "Line 1" & _ \n "Line 2" & _ ...`
   - **Good:** Use `s = s & "Line X" & vbCrLf` for every line.
3. **Ghost Forms (Error 75):** If you delete a generated form and immediately run the builder script again to generate a form with the *same name*, VBE might throw `Run-time error '75': Path/File access error`. This is because VBE locks the temporary `.frm` file in your `%TEMP%` folder. To fix this, simply close and restart CorelDRAW.

## Boilerplate Template

Save this as a `.bas` file and use it as a starting point for future macro forms:

```vb
Attribute VB_Name = "mdlFormBuilder_Template"
Option Explicit

' Auto-generates a UserForm and injects code-behind.
Public Sub BuildMyForm()
    Dim vbProj As Object
    Set vbProj = Application.VBE.ActiveVBProject
    
    ' 1. Remove existing form to prevent duplicates
    On Error Resume Next
    Dim existingComp As Object
    Set existingComp = vbProj.VBComponents("frmMyNewForm")
    If Not existingComp Is Nothing Then
        vbProj.VBComponents.Remove existingComp
    End If
    On Error GoTo 0
    
    ' 2. Add the UserForm (3 = MSForm)
    Dim newFormComp As Object
    Set newFormComp = vbProj.VBComponents.Add(3)
    newFormComp.Name = "frmMyNewForm"
    
    ' 3. Configure the Form
    With newFormComp.Properties
        .Item("Caption").Value = "My Auto-Generated Form"
        .Item("Width").Value = 300
        .Item("Height").Value = 200
    End With
    
    Dim formDesigner As Object
    Set formDesigner = newFormComp.Designer
    
    ' 4. Add UI Controls
    Dim cmdBtn As Object
    Set cmdBtn = formDesigner.Controls.Add("Forms.CommandButton.1", "cmdAction")
    With cmdBtn
        .Caption = "Click Me"
        .Left = 100: .Top = 100: .Width = 80: .Height = 24
    End With

    ' 5. Inject Code-Behind
    InjectCodeBehind newFormComp.CodeModule
    MsgBox "Form built successfully!", vbInformation
End Sub

Private Sub InjectCodeBehind(ByVal codeMod As Object)
    Dim s As String
    ' Always append iteratively to avoid the 24 line-continuation limit!
    s = "Option Explicit" & vbCrLf
    s = s & "Private Sub cmdAction_Click()" & vbCrLf
    s = s & "    MsgBox ""Hello from the auto-generated form!""" & vbCrLf
    s = s & "End Sub" & vbCrLf

    codeMod.DeleteLines 1, codeMod.CountOfLines
    codeMod.AddFromString s
End Sub
```
