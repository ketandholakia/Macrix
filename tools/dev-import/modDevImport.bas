Attribute VB_Name = "modDevImport"
Option Explicit
''
' modDevImport — developer bootstrap importer (dev tool; NOT a user feature).
'
' Purpose: pull every staged .bas/.frm from a staging folder into the VBA project
' this module lives in, so a source change on disk can be pushed into the live
' Macrix.gms with one call — after a ONE-TIME manual import of this file.
'
' Why a separate file: this module lives in tools\dev-import and is deliberately
' NOT part of the staged set, so it never removes/re-imports itself mid-run.
'
' PREREQUISITE (verified against the CorelDRAW 2021 type library: IVGApplication
' exposes VBE, GMSManager, InitializeVBA):
'   CorelDRAW must allow access to the VBA project object model. If not,
'   `Application.VBE` is unavailable and this module reports that clearly.
'   In CorelDRAW: Tools > Options > VBA (or the VBE's own security options).
'
' Depends on: nothing (standalone by design — it may run before the framework
' modules are present).
' CorelDRAW: yes (VBIDE automation) — characterization pending.

Private Const Dev_SelfName As String = "modDevImport"
Private Const Dev_StagingEnv As String = "MACRIX_STAGING"
Private Const Dev_StagingSub As String = "Macrix\staging"

' Import (or replace) every staged module into this project.
Public Sub DevImport_RunAll()
    On Error GoTo Cleanup
    Dim vbe As Object, proj As Object
    Dim dir As String

    Set vbe = p_VBE()
    If vbe Is Nothing Then
        MsgBox "Cannot access the VBA project object model." & vbCrLf & vbCrLf & _
               "CorelDRAW must allow VBA project access. Enable it in" & vbCrLf & _
               "Tools > Options > VBA (CorelDRAW) / the VBE security options," & vbCrLf & _
               "then restart CorelDRAW and try again.", _
               vbExclamation, "Macrix Dev Import"
        GoTo Cleanup
    End If

    Set proj = p_FindOwnProject(vbe)
    If proj Is Nothing Then
        MsgBox "Could not find the VBA project that contains " & Dev_SelfName & ".", _
               vbCritical, "Macrix Dev Import"
        GoTo Cleanup
    End If

    dir = p_StagingDir()
    If Dir$(dir, vbDirectory) = "" Then
        MsgBox "Staging folder not found:" & vbCrLf & dir & vbCrLf & vbCrLf & _
               "Run  build\deploy.ps1 -Stage  first to stage modules.", _
               vbExclamation, "Macrix Dev Import"
        GoTo Cleanup
    End If

    Dim imported As Long, replaced As Long, failed As Long
    p_ImportDir proj, dir, imported, replaced, failed

    MsgBox "Dev import complete." & vbCrLf & _
           "  Project : " & CStr(proj.Name) & vbCrLf & _
           "  Imported: " & CStr(imported) & vbCrLf & _
           "  Replaced: " & CStr(replaced) & vbCrLf & _
           "  Failed  : " & CStr(failed) & vbCrLf & vbCrLf & _
           "From: " & dir, vbInformation, "Macrix Dev Import"

Cleanup:
    If Err.Number <> 0 Then
        MsgBox "Dev import failed: [" & CStr(Err.Number) & "] " & Err.Description, _
               vbCritical, "Macrix Dev Import"
    End If
End Sub

' ---- private helpers ----

Private Function p_VBE() As Object
    On Error Resume Next
    Set p_VBE = Application.VBE
    On Error GoTo 0
End Function

' The project owning this module is the target (so the caller does not have to
' name it). Falls back to Nothing when VBE is not enumerable.
Private Function p_FindOwnProject(ByVal vbe As Object) As Object
    On Error Resume Next
    Dim proj As Object, comp As Object
    For Each proj In vbe.VBProjects
        For Each comp In proj.VBComponents
            If StrComp(comp.Name, Dev_SelfName, vbTextCompare) = 0 Then
                Set p_FindOwnProject = proj
                Exit Function
            End If
        Next comp
    Next proj
    On Error GoTo 0
End Function

Private Function p_StagingDir() As String
    Dim base As String
    base = Environ$(Dev_StagingEnv)
    If base <> "" Then
        p_StagingDir = base
        Exit Function
    End If
    base = Environ$("APPDATA")
    If base = "" Then base = Environ$("USERPROFILE")
    p_StagingDir = base & "\" & Dev_StagingSub
End Function

Private Sub p_ImportDir(ByVal proj As Object, ByVal dir As String, _
                        ByRef imported As Long, ByRef replaced As Long, ByRef failed As Long)
    Dim f As String, base As String, path As String
    Dim comp As Object

    f = Dir$(dir & "\*.bas")
    Do While Len(f) > 0
        path = dir & "\" & f
        base = p_BaseName(f)
        If StrComp(base, Dev_SelfName, vbTextCompare) <> 0 Then
            Set comp = p_FindComponent(proj, base)
            If Not comp Is Nothing Then
                On Error Resume Next
                proj.VBComponents.Remove comp
                On Error GoTo 0
                replaced = replaced + 1
            End If
            On Error Resume Next
            proj.VBComponents.Import path
            If Err.Number = 0 Then
                imported = imported + 1
            Else
                failed = failed + 1
            End If
            On Error GoTo 0
        End If
        f = Dir$()
    Loop

    f = Dir$(dir & "\*.frm")
    Do While Len(f) > 0
        path = dir & "\" & f
        base = p_BaseName(f)
        Set comp = p_FindComponent(proj, base)
        If Not comp Is Nothing Then
            On Error Resume Next
            proj.VBComponents.Remove comp
            On Error GoTo 0
            replaced = replaced + 1
        End If
        On Error Resume Next
        proj.VBComponents.Import path
        If Err.Number = 0 Then
            imported = imported + 1
        Else
            failed = failed + 1
        End If
        On Error GoTo 0
        f = Dir$()
    Loop
End Sub

Private Function p_FindComponent(ByVal proj As Object, ByVal name As String) As Object
    On Error Resume Next
    Dim comp As Object
    For Each comp In proj.VBComponents
        If StrComp(comp.Name, name, vbTextCompare) = 0 Then
            Set p_FindComponent = comp
            Exit Function
        End If
    Next comp
    On Error GoTo 0
End Function

' Strip the extension from a file name (no FileSystemObject dependency).
Private Function p_BaseName(ByVal fileName As String) As String
    Dim i As Long
    For i = Len(fileName) To 1 Step -1
        If Mid$(fileName, i, 1) = "." Then
            p_BaseName = Left$(fileName, i - 1)
            Exit Function
        End If
    Next i
    p_BaseName = fileName
End Function
