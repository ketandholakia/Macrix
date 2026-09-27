Attribute VB_Name = "modGangJob"
'==============================================================
' modGangJob
' Phase 1 of the Gang Job macro: an in-memory data model of jobs
' to be ganged, plus persistence of each job's Quantity/Name/
' AllowRotation onto the artwork itself (via CorelDRAW's Shape
' Properties collection), so the list can be rebuilt later without
' re-typing anything. Width/Height are NEVER persisted this way --
' they are always re-measured from the artwork's current bounding
' box, so a resized job can't silently go stale.
'
' Metadata mechanism (see CorelDRAW SDK article "Storing Custom
' Information in Documents"): Shape.Properties(SolutionID, index).
' Chosen over Object Data Manager (Shape.ObjectData) because it is
' hidden from the user and not limited to a fixed set of visible
' fields. GANGJOB_SOLUTION_ID below is this project's unique GUID
' for that mechanism -- do not change it once jobs have been tagged
' in real documents, or previously-tagged artwork becomes unreadable
' (Properties.Exists will simply return False for the new GUID).
'
' Properties are per-Shape, not per-ShapeRange. A job's artwork can
' be several ungrouped shapes (same "treat the selection as one
' unit" approach as modTileFill), so every member shape of the
' selection is tagged with the SAME JobID (index 4) -- that is what
' lets ScanDocumentForGangJobs re-assemble the whole job (and its
' combined bounding box) from possibly-multiple tagged shapes.
'
' Workflow:
'   1. Select the artwork for one job in CorelDRAW.
'   2. Run GangJob_AddFromSelection (in modMain) -- reads the
'      selection's bounding box automatically, prompts for Name/
'      Quantity/rotation, and tags every shape in the selection.
'   3. Repeat for each job. GangJob_ShowJobs dumps the in-memory
'      list to the Immediate Window (Ctrl+G) to verify it.
'   4. Later (e.g. after reopening the file), run
'      GangJob_ScanDocument to rebuild the list purely from tagged
'      artwork already in the document -- no re-entry needed.
'
' Later phases add: modGangJobOptimizer (candidate generation +
' scoring, working only with clsGangJob's numbers -- no CorelDRAW
' objects), modGangJobOutput (CorelDRAW page creation), and a real
' frmGangJobSettings UI replacing the InputBox prompts below.
'==============================================================

Option Explicit
Option Private Module

' This project's unique ID for the Properties mechanism. Generated
' once; keep stable (see header note above).
Private Const GANGJOB_SOLUTION_ID As String = "751F5091-2D74-4614-8CA0-012C7FF190A6"

' Property indices under GANGJOB_SOLUTION_ID:
Private Const PROP_JOBNAME As Long = 1        ' String
Private Const PROP_QUANTITY As Long = 2       ' Long
Private Const PROP_ALLOWROTATE As Long = 3    ' Long, 0/1 (Properties has no native Boolean)
Private Const PROP_JOBID As Long = 4          ' Long -- ties multiple shapes to one job

Public g_GangJobs As Collection

'--------------------------------------------------------------
' Reads the current selection's bounding box and adds it as a new
' job, after prompting for a name, a required quantity, and whether
' 90-degree rotation is allowed for this job. Tags every shape in
' the selection with the same job metadata so ScanDocumentForGangJobs
' can find it again later.
'--------------------------------------------------------------
Public Function AddJobFromSelection() As Boolean

    Dim doc As Document
    Dim sel As ShapeRange
    Dim selCount As Long

    If Not g_GangSettingsInitialized Then InitDefaultGangSettings
    If g_GangJobs Is Nothing Then Set g_GangJobs = New Collection

    If ActiveDocument Is Nothing Then
        MsgBox "Open a document first.", vbExclamation
        Exit Function
    End If

    Set doc = ActiveDocument
    doc.Unit = g_GangUnit

    ' Guard against calling .Count on a Nothing SelectionRange -- same
    ' two-step check modTileFill.RunTileFill uses, since VBA's "Or"
    ' does not short-circuit.
    selCount = 0
    On Error Resume Next
    selCount = doc.SelectionRange.Count
    On Error GoTo 0
    If selCount = 0 Then
        MsgBox "Select the artwork for this job first.", vbExclamation
        Exit Function
    End If

    Set sel = doc.SelectionRange

    Dim w As Double, h As Double
    w = sel.SizeWidth
    h = sel.SizeHeight
    If w <= 0 Or h <= 0 Then
        MsgBox "Could not read the selection's size.", vbExclamation
        Exit Function
    End If

    Dim jobName As String
    jobName = InputBox("Job name:", "Add Gang Job", "Job " & (g_GangJobs.Count + 1))
    If jobName = "" Then Exit Function

    Dim qtyResp As String
    qtyResp = InputBox("Quantity required:", "Add Gang Job", "1")
    If qtyResp = "" Then Exit Function
    If Not IsNumeric(qtyResp) Or CLng(Val(qtyResp)) <= 0 Then
        MsgBox "Quantity must be a positive number.", vbExclamation
        Exit Function
    End If

    Dim rotResp As String
    rotResp = InputBox("Allow 90-degree rotation for this job? (Y/N):", _
                        "Add Gang Job", "Y")
    If rotResp = "" Then Exit Function

    Dim job As clsGangJob
    Set job = New clsGangJob
    job.JobID = g_GangNextJobID
    job.JobName = jobName
    job.Width = w
    job.Height = h
    job.Quantity = CLng(Val(qtyResp))
    job.AllowRotation = (UCase(Left(rotResp, 1)) = "Y")

    g_GangJobs.Add job, CStr(job.JobID)
    g_GangNextJobID = g_GangNextJobID + 1

    TagShapeRangeWithJob sel, job

    MsgBox "Added job """ & job.JobName & """: " & _
           Format(job.Width, "0.###") & " x " & Format(job.Height, "0.###") & _
           " (" & UnitName(g_GangUnit) & "), qty " & job.Quantity & ".", vbInformation

    AddJobFromSelection = True

End Function

'--------------------------------------------------------------
' Writes job metadata onto every shape in the range so the job can
' be found again by ScanDocumentForGangJobs, even if its artwork is
' several ungrouped shapes.
'--------------------------------------------------------------
Private Sub TagShapeRangeWithJob(sr As ShapeRange, job As clsGangJob)
    Dim s As Shape
    For Each s In sr.Shapes
        s.Properties(GANGJOB_SOLUTION_ID, PROP_JOBNAME) = job.JobName
        s.Properties(GANGJOB_SOLUTION_ID, PROP_QUANTITY) = job.Quantity
        s.Properties(GANGJOB_SOLUTION_ID, PROP_ALLOWROTATE) = IIf(job.AllowRotation, 1, 0)
        s.Properties(GANGJOB_SOLUTION_ID, PROP_JOBID) = job.JobID
    Next s
End Sub

'--------------------------------------------------------------
' Rebuilds g_GangJobs entirely from Properties-tagged shapes found
' in the active document. Shapes sharing the same PROP_JOBID are
' grouped into one job, and that job's Width/Height are re-measured
' from their COMBINED current bounding box (never trusted from the
' stored properties -- there aren't any; only Name/Quantity/Rotate/
' JobID are stored).
'
' This REPLACES the current in-memory list. Run it to pick up a
' document's tagged jobs fresh (e.g. after reopening a file) -- not
' to merge with jobs already added this session.
'--------------------------------------------------------------
Public Function ScanDocumentForGangJobs() As Boolean

    If ActiveDocument Is Nothing Then
        MsgBox "Open a document first.", vbExclamation
        Exit Function
    End If

    If Not g_GangSettingsInitialized Then InitDefaultGangSettings

    Dim doc As Document
    Set doc = ActiveDocument
    doc.Unit = g_GangUnit

    ' Late-bound Dictionary (no project reference needed, per this
    ' project's late-bound convention): JobID (as String key) ->
    ' ShapeRange of every shape tagged with that JobID, across all
    ' pages.
    Dim ranges As Object
    Set ranges = CreateObject("Scripting.Dictionary")

    Dim pg As Page, s As Shape
    For Each pg In doc.Pages
        For Each s In pg.Shapes.All
            If s.Properties.Exists(GANGJOB_SOLUTION_ID, PROP_JOBID) Then
                Dim key As String
                key = CStr(CLng(s.Properties(GANGJOB_SOLUTION_ID, PROP_JOBID)))

                Dim sr As ShapeRange
                If ranges.Exists(key) Then
                    Set sr = ranges(key)
                Else
                    Set sr = New ShapeRange
                    ranges.Add key, sr
                End If
                sr.Add s
            End If
        Next s
    Next pg

    If ranges.Count = 0 Then
        MsgBox "No gang-job-tagged artwork found in this document.", vbInformation
        Exit Function
    End If

    Set g_GangJobs = New Collection
    g_GangNextJobID = 1

    Dim k As Variant
    For Each k In ranges.Keys
        Set sr = ranges(k)

        Dim first As Shape
        Set first = sr.Shapes(1)

        Dim job As clsGangJob
        Set job = New clsGangJob
        job.JobID = CLng(k)
        job.JobName = CStr(first.Properties(GANGJOB_SOLUTION_ID, PROP_JOBNAME))
        job.Quantity = CLng(first.Properties(GANGJOB_SOLUTION_ID, PROP_QUANTITY))
        job.AllowRotation = (CLng(first.Properties(GANGJOB_SOLUTION_ID, PROP_ALLOWROTATE)) <> 0)
        job.Width = sr.SizeWidth
        job.Height = sr.SizeHeight

        g_GangJobs.Add job, CStr(job.JobID)
        If job.JobID >= g_GangNextJobID Then g_GangNextJobID = job.JobID + 1
    Next k

    MsgBox "Rebuilt " & g_GangJobs.Count & " job(s) from tagged artwork.", vbInformation

    ScanDocumentForGangJobs = True

End Function

'--------------------------------------------------------------
' Removes one job by its JobID (not its position -- positions shift
' as jobs are added/removed, IDs don't). Does NOT remove the
' Properties tags from the shapes themselves -- re-run
' ScanDocumentForGangJobs and this job will reappear. Untag
' manually if that's not wanted (not yet implemented).
'--------------------------------------------------------------
Public Sub RemoveGangJob(jobID As Long)
    If g_GangJobs Is Nothing Then Exit Sub
    On Error Resume Next
    g_GangJobs.Remove CStr(jobID)
    On Error GoTo 0
End Sub

Public Sub ClearGangJobs()
    Set g_GangJobs = New Collection
End Sub

'--------------------------------------------------------------
' Prints the current job list to the Immediate Window -- the Phase
' 1 stand-in for a real list UI.
'--------------------------------------------------------------
Public Sub DumpGangJobs()
    If g_GangJobs Is Nothing Then
        Debug.Print "GangJobs: (empty)"
        Exit Sub
    End If
    If g_GangJobs.Count = 0 Then
        Debug.Print "GangJobs: (empty)"
        Exit Sub
    End If

    Dim j As clsGangJob
    Debug.Print "GangJobs (" & g_GangJobs.Count & "):"
    For Each j In g_GangJobs
        Debug.Print "  #" & j.JobID & " """ & j.JobName & """  " & _
                     Format(j.Width, "0.###") & " x " & Format(j.Height, "0.###") & _
                     "  qty=" & j.Quantity & _
                     "  rotate=" & j.AllowRotation
    Next j
End Sub

Private Function UnitName(u As cdrUnit) As String
    Select Case u
        Case cdrMillimeter: UnitName = "mm"
        Case cdrInch: UnitName = "inch"
        Case cdrPoint: UnitName = "pt"
        Case cdrPixel: UnitName = "px"
        Case Else: UnitName = "units"
    End Select
End Function
