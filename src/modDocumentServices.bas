Attribute VB_Name = "modDocumentServices"
Option Explicit
''
' modDocumentServices — Document-level manipulation. Late-bound.
' Depends on: modComLifecycle. CorelDRAW: yes (characterization).

Public Function SV_Doc_Active(ByVal app As Object) As Object
    Set SV_Doc_Active = modAppServices.SV_App_ActiveDocument(app)
End Function

Public Function SV_Doc_FullName(ByVal doc As Object) As String
    On Error Resume Next
    SV_Doc_FullName = CStr(CallByName(doc, "FullFileName", VbGet))
    On Error GoTo 0
End Function

Public Function SV_Doc_IsDirty(ByVal doc As Object) As Boolean
    On Error Resume Next
    SV_Doc_IsDirty = CBool(CallByName(doc, "Dirty", VbGet))
    On Error GoTo 0
End Function

' Save the document; returns True if it changed name (Save As) else False.
Public Function SV_Doc_Save(ByVal doc As Object, Optional ByVal path As String = "") As Boolean
    On Error GoTo Fail
    If path = "" Then
        CallByName doc, "Save", VbMethod
    Else
        CallByName doc, "SaveAs", VbMethod, path
    End If
    SV_Doc_Save = True
    Exit Function
Fail:
    modLogger.LOG_Error "SV_Doc_Save failed (" & CStr(Err.Number) & "): " & Err.Description
    SV_Doc_Save = False
End Function