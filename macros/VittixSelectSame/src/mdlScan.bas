Attribute VB_Name = "mdlScan"
Option Explicit

' ===========================================================================
' mdlScan
'
' Builds the candidate shape list according to the chosen search scope, then
' (later, via SelectMatches) turns a list of matches into an actual
' CorelDRAW selection. Traversal is iterative-recursive over groups so that
' "Search Inside Groups" can descend into arbitrarily nested groups without
' altering the document (no ungrouping, no curve conversion).
' ===========================================================================

Public Function CollectCandidates(opts As TScanOptions) As Collection
    Dim results As New Collection

    Select Case opts.Scope
        Case SCOPE_CURRENT_PAGE
            ScanShapeCollection ActiveDocument.ActivePage.shapes, opts, results

        Case SCOPE_ALL_PAGES
            Dim pg As Page
            For Each pg In ActiveDocument.Pages
                ScanShapeCollection pg.shapes, opts, results
            Next pg

        Case SCOPE_CURRENT_LAYER
            ScanShapeCollection ActiveLayer.shapes, opts, results

        Case SCOPE_ALL_LAYERS
            Dim ly As Layer
            For Each ly In ActiveDocument.ActivePage.Layers
                ScanShapeCollection ly.shapes, opts, results
            Next ly

        Case SCOPE_CURRENT_SELECTION
            Dim sr As ShapeRange
            Set sr = ActiveDocument.ActiveSelectionRange
            Dim i As Long
            For i = 1 To sr.Count
                AddShapeOrRecurse sr(i), opts, results
            Next i
    End Select

    Set CollectCandidates = results
End Function

Private Sub ScanShapeCollection(shapes As shapes, opts As TScanOptions, results As Collection)
    Dim sh As Shape
    On Error Resume Next
    For Each sh In shapes
        AddShapeOrRecurse sh, opts, results
    Next sh
    On Error GoTo 0
End Sub

' Kept close to the naming used in the spec (ScanPage / ScanLayer / ScanGroup
' / ScanShape) as thin wrappers, in case you want to call an individual
' scope directly from other code:
Public Sub ScanPage(pg As Page, opts As TScanOptions, results As Collection)
    ScanShapeCollection pg.shapes, opts, results
End Sub

Public Sub ScanLayer(ly As Layer, opts As TScanOptions, results As Collection)
    ScanShapeCollection ly.shapes, opts, results
End Sub

Public Sub ScanGroup(grp As Shape, opts As TScanOptions, results As Collection)
    ScanShapeCollection grp.shapes, opts, results
End Sub

Public Sub ScanShape(sh As Shape, opts As TScanOptions, results As Collection)
    AddShapeOrRecurse sh, opts, results
End Sub

Private Sub AddShapeOrRecurse(sh As Shape, opts As TScanOptions, results As Collection)
    On Error Resume Next
    If sh.Type = cdrGroupShape Then
        If opts.SearchInsideGroups Then
            ScanGroup sh, opts, results
            Exit Sub
        End If
        ' Search Inside Groups is off: the group itself is a normal
        ' candidate (e.g. can match on ObjectType = Group), its children
        ' are not inspected.
    End If
    AddCandidateIfEligible sh, opts, results
End Sub

Private Sub AddCandidateIfEligible(sh As Shape, opts As TScanOptions, results As Collection)
    On Error Resume Next
    If Not opts.IncludeHidden And Not sh.Visible Then Exit Sub
    If Not opts.IncludeLocked And sh.Locked Then Exit Sub
    results.Add sh
End Sub

' ---------------------------------------------------------------------------
' SelectMatches - applies the final selection in CorelDRAW.
'
' Locked objects: CorelDRAW will not let a locked shape join a selection.
' If "Include Locked Objects" was checked, matching locked shapes are
' unlocked so they can be selected - there is no API to select-while-locked,
' so this is a documented, unavoidable side effect (the caller is told how
' many objects were affected so it can inform the user).
' ---------------------------------------------------------------------------
Public Function SelectMatches(matches As Collection, refShape As Shape, opts As TScanOptions) As Long
    Dim sr As ShapeRange
    Set sr = ActiveDocument.CreateShapeRange

    Dim unlockedCount As Long
    unlockedCount = 0

    Dim sh As Shape

    If opts.IncludeReference Then
        On Error Resume Next
        sr.Add refShape
        On Error GoTo 0
    End If

    For Each sh In matches
        On Error Resume Next
        If opts.IncludeLocked And sh.Locked Then
            sh.Locked = False
            unlockedCount = unlockedCount + 1
        End If
        sr.Add sh
        On Error GoTo 0
    Next sh

    If opts.AddToExisting Then
        For Each sh In sr
            On Error Resume Next
            sh.AddToSelection
            On Error GoTo 0
        Next sh
    Else
        On Error Resume Next
        sr.CreateSelection
        On Error GoTo 0
    End If

    SelectMatches = unlockedCount
End Function

