Attribute VB_Name = "feat_ExportText"
Option Explicit
''
' feat_ExportText — sample feature: write every text shape's content, across all
' pages of the active document, to a timestamped .txt file under the user data
' directory. Demonstrates the safe-run entry pattern and path management
' (no hard-coded paths).
'
' Text extraction (Pages/Shapes/Text/Story member names) is Corel-dependent and
' version-sensitive — characterization pending (§40), see
' tests/manual/CHARACTERIZATION.md. File IO itself is pure.
'
' Depends on: modAppServices, modPathManager, modExportImportServices,
' modBinding, modComLifecycle, modUiUtilities, modErrorHandler.
' CorelDRAW: yes (characterization).

Private Const F_BreakerId As String = "export-text"

Public Sub Main()
    Dim app As Object, doc As Object
    On Error GoTo Cleanup
    If modErrorHandler.EH_IsDisabled(F_BreakerId) Then
        Err.Raise 2001, "feat_ExportText.Main", "feature session-disabled"
    End If
    Set app = modAppServices.SV_App_Host()
    If app Is Nothing Then Err.Raise 2002, "feat_ExportText.Main", "host unavailable"
    Set doc = modAppServices.SV_App_ActiveDocument(app)
    If doc Is Nothing Then Err.Raise 2003, "feat_ExportText.Main", "no active document"

    Dim outFile As String
    outFile = modExportImportServices.SVE_BuildExportPath( _
        modPathManager.PF_GetUserDataDir(), "document_text", ".txt")
    p_WriteDocumentText doc, outFile
    modErrorHandler.EH_RecordSuccess F_BreakerId
    modUiUtilities.UU_MsgInfo "Text exported to " & outFile

Cleanup:
    modComLifecycle.CM_Release app
    modComLifecycle.CM_Release doc
    If Err.Number <> 0 Then
        modErrorHandler.EH_ReportError Err.Number, "feat_ExportText.Main", Err.Description
        If modErrorHandler.EH_RecordFailure(F_BreakerId) Then
            modUiUtilities.UU_MsgWarn "This feature is disabled for this session; see the log."
        End If
    End If
End Sub

' Enumerate every shape on every page of `doc` and write each text shape's
' Story (via Shape.Text.Story) to `path`, one per line. Late-bound and guarded
' member-by-member (§42) since Pages/Shapes/Text/Story availability varies by
' CorelDRAW version — confirm against the installed type library before
' promoting this feature out of beta.
'
' Unlike most service helpers in this codebase, real file-IO failures (Open/
' Print/Close) are NOT swallowed here — they propagate to Main's Cleanup label
' so a failed export is reported and counted by the circuit breaker instead of
' silently claiming success.
Private Sub p_WriteDocumentText(ByVal doc As Object, ByVal path As String)
    Dim h As Integer
    h = FreeFile
    On Error GoTo Fail
    Open path For Output As #h

    Dim pages As Object, pg As Object, shapes As Object, sh As Object, txt As Object
    Dim i As Long, j As Long, pageCount As Long, shapeCount As Long
    Dim story As String

    Set pages = CallByName(doc, "Pages", VbGet)
    pageCount = CLng(CallByName(pages, "Count", VbGet))
    For i = 1 To pageCount
        Set pg = CallByName(pages, "Item", VbMethod, i)
        Set shapes = CallByName(pg, "Shapes", VbGet)
        shapeCount = CLng(CallByName(shapes, "Count", VbGet))
        For j = 1 To shapeCount
            Set sh = CallByName(shapes, "Item", VbMethod, j)
            If modBinding.BND_HasMember(sh, "Text") Then
                On Error Resume Next
                story = ""
                Set txt = CallByName(sh, "Text", VbGet)
                If Not txt Is Nothing Then story = CStr(CallByName(txt, "Story", VbGet))
                On Error GoTo Fail
                If Len(story) > 0 Then Print #h, story
                modComLifecycle.CM_Release txt
            End If
            modComLifecycle.CM_Release sh
        Next j
        modComLifecycle.CM_Release shapes
        modComLifecycle.CM_Release pg
    Next i
    modComLifecycle.CM_Release pages

    Close #h
    Exit Sub

Fail:
    Dim savedNum As Long, savedDesc As String
    savedNum = Err.Number
    savedDesc = Err.Description
    On Error Resume Next
    Close #h
    On Error GoTo 0
    Err.Raise savedNum, "feat_ExportText.p_WriteDocumentText", savedDesc
End Sub
