Attribute VB_Name = "modImposition"
'==============================================================
' modImposition
' Core imposition logic for CorelDRAW: Simple N-up, Step & Repeat,
' and Signature (saddle-stitch booklet) layouts.
'
' TWO WAYS TO RUN:
'   1. ShowImpositionForm  -> opens frmImpositionSettings so you can
'                             set everything interactively, then runs.
'   2. RunImposition       -> runs directly using whatever is in
'                             modSettings (defaults, if the form has
'                             never been opened this session).
'
' Settings live in modSettings.bas as Public "g_..." variables —
' see that file to change defaults.
'
' NOTE: CorelDRAW's VBA object model varies slightly between
' versions (X7, 2018, 2021, 2024...). If a line errors, it is
' almost always SetPosition / Layer assignment / CreateLineSegment
' — check the equivalent method name in your version's VBA object
' browser (F2 in the VBA IDE, search "Shape" or "Layer").
'
' 2026-09 FIX (z-order): PlacePageInCell picked out "the shapes this
' call just copied" by assuming they land at index beforeCount+1..Count
' in Page.Shapes. CorelDRAW actually inserts freshly copied shapes at
' the FRONT of the z-order (new Shapes(1)), pushing existing shapes to
' higher indices. So every call after the first re-grabbed and re-moved
' the PREVIOUS page's already-placed shapes, while the current page's
' real new shapes sat untouched at raw source coordinates -- all pages
' piled on top of each other instead of spreading across the grid. Now
' identified by stable Shape.StaticID instead of position.
'
' 2026-09 FIX (placement / sheet authority / origin):
'   (1) RunImposition NO LONGER overwrites g_SheetWidth/g_SheetHeight
'       from the source page size. The user-entered sheet size is
'       authoritative (e.g. 450 x 320 mm stays 450 x 320 mm). Cells are
'       the source trim at 1:1 (no scaling); if the grid does not fit
'       the sheet, ValidateSheetFits reports required vs available and
'       aborts instead of silently pushing pages off-sheet.
'   (2) Cell coordinates are derived from the ACTUAL output page origin
'       (outPage.LeftX / TopY / BottomY), not from a hard-coded (0,0):
'         cellLeft   = outPage.LeftX + g_MarginLeft + c*(cellW+g_GutterX)
'         cellBottom = outPage.TopY - g_MarginTop - cellH - r*(cellH+g_GutterY)
'       CorelDRAW Y grows upward (origin at lower-left by default), so
'       Row 0 lands at the top. This holds for any page origin, not just
'       LeftX=0/BottomY=0 (see Test 5: non-default origin).
'   (3) PlacePageInCell uses the ACTUAL post-paste bounding box, not an
'       assumed source origin and not a tiny origin marker. The old marker
'       (CreateRectangle at assumed origin + PositionX/Y correction) was
'       unreliable: CopyToLayer preserves each source page's absolute
'       document coordinates, so pages copied from different source pages
'       land at different absolute baselines, and a marker created with
'       CreateRectangle(0,0,...) does not represent srcPage.LeftX/BottomY.
'       That is why Page 4 (first paste, empty sheet) landed correctly
'       while Page 1 (second paste, different source baseline) was shifted
'       right/off-sheet. Now: measure source-content bbox before copy
'       (to preserve the content-to-trim offset), measure pasted union bbox
'       after CopyToLayer, then Move(desired - actual). No marker, no
'       PositionX/Y, no hard-coded offsets. Move(dx,dy) is
'       reference-point independent.
'   (4) Two-pass build: ALL sheets are AddPages+SetSize first (Pass A),
'       content+marks placed afterwards (Pass B) from final frames. This
'       avoids stranding already-placed content when a later AddPages
'       shifts existing page frames, and avoids creating anything on
'       source pages (which itself can shift that page's frame).
'   (5) DrawCropMarks uses the same absolute cell corners so marks stay
'       with their cell; each created segment is re-measured and nudged
'       by Move so a local-vs-absolute creation interpretation cannot
'       detach marks from the cell. Marks may extend into margin/bleed
'       (preserved behavior) but never move the page itself off-sheet.
'==============================================================

Option Explicit
Option Private Module

' Returns True only when the imposition completed successfully, so the
' caller (frmImpositionSettings.cmdRun_Click) can persist last-used
' settings exclusively on success. Any early exit leaves False.
Public Function RunImposition() As Boolean

    Dim srcDoc As Document
    Dim outDoc As Document
    Dim totalPages As Integer
    Dim pagesPerSheet As Integer
    Dim totalSheets As Integer
    Dim sheetIndex As Integer
    Dim r As Integer, c As Integer
    Dim outPage As Page
    Dim cellW As Double, cellH As Double
    Dim orderArr() As Integer
    Dim orderLen As Integer
    Dim slotIndex As Integer

    RunImposition = False

    If Not g_SettingsInitialized Then InitDefaultSettings

    If ActiveDocument Is Nothing Then
        MsgBox "Open a document first.", vbExclamation
        Exit Function
    End If

    Set srcDoc = ActiveDocument
    srcDoc.Unit = g_Unit
    totalPages = srcDoc.Pages.Count

    If totalPages = 0 Then
        MsgBox "The active document has no pages.", vbExclamation
        Exit Function
    End If

    pagesPerSheet = g_GridRows * g_GridCols

    ' Build the sequence of source-page indices to drop into each
    ' cell, slot by slot (0 = leave this slot blank).
    If Not BuildPageOrder(totalPages, pagesPerSheet, orderArr) Then
        Exit Function ' BuildPageOrder already showed an error message
    End If
    orderLen = UBound(orderArr) - LBound(orderArr) + 1
    totalSheets = Int((orderLen - 1) / pagesPerSheet) + 1

    ' Cell = source trim at 1:1, NO scaling. Sheet size stays exactly as
    ' the user entered it (authoritative). Do NOT recompute g_SheetWidth/
    ' g_SheetHeight here -- that was the main placement bug: it silently
    ' replaced e.g. 450x320 with 446.8x289.4 and hid fit errors.
    cellW = srcDoc.Pages(1).SizeWidth
    cellH = srcDoc.Pages(1).SizeHeight

    Debug.Print "modImposition: src trim cell=" & Round(cellW, 3) & " x " & Round(cellH, 3) & _
                " sheet(requested)=" & Round(g_SheetWidth, 3) & " x " & Round(g_SheetHeight, 3) & _
                " grid=" & g_GridRows & "x" & g_GridCols & _
                " gutter=" & Round(g_GutterX, 3) & "," & Round(g_GutterY, 3) & _
                " marginLT=" & Round(g_MarginLeft, 3) & "," & Round(g_MarginTop, 3)

    ' All source pages are expected to share the trim size (typical). If a
    ' later page is larger than the cell, placing at 100% would overflow,
    ' so fail loudly instead of scaling or pushing off-sheet.
    Dim pCheck As Integer
    For pCheck = 1 To totalPages
        If Abs(srcDoc.Pages(pCheck).SizeWidth - cellW) > 0.01 Or _
           Abs(srcDoc.Pages(pCheck).SizeHeight - cellH) > 0.01 Then
            MsgBox "Source pages have different sizes." & vbCrLf & _
                   "Page 1: " & Round(cellW, 2) & " x " & Round(cellH, 2) & vbCrLf & _
                   "Page " & pCheck & ": " & _
                   Round(srcDoc.Pages(pCheck).SizeWidth, 2) & " x " & _
                   Round(srcDoc.Pages(pCheck).SizeHeight, 2) & vbCrLf & _
                   "Imposition places at 100% with no scaling, so mixed sizes need identical cells.", _
                   vbExclamation
            Exit Function
        End If
    Next pCheck

    ' Authoritative fit check: trim + gutters + symmetric margins vs sheet.
    ' Bleed / crop length are deliberately NOT part of trim fit (they may
    ' legally extend into margin area -- preserved behavior).
    If Not ValidateSheetFits(cellW, cellH) Then Exit Function

    Application.Optimization = True
    Application.EventsEnabled = False

    ' Create a fresh output document
    ' Build the sheets inside a DUPLICATE of the source document.
    ' Every cross-document route is blocked in this object model: the clipboard carries
    ' nothing, CopyToLayer refuses ("Specified object is from another document") and
    ' Document.Export fails with a type mismatch (13). CopyToLayer DOES work within one
    ' document, so the source pages come along in the duplicate and are deleted again once
    ' the sheets have been built.
    Set outDoc = srcDoc.Duplicate
    outDoc.Unit = g_Unit
    outDoc.ReferencePoint = cdrTopLeft

    Application.Optimization = False
    Application.EventsEnabled = True

    ' Two passes (live-probe verified on CorelDRAW 2021): AddPages/SetSize
    ' can shift EXISTING pages' absolute frames (sources moved 0,0 -> offset
    ' when the first big sheet was sized; creating shapes on a source page
    ' can shift that page again). Pass A creates+sizes ALL sheets first, so
    ' layout settles with no placed content to strand. Pass B only reads
    ' final frames and pastes -- never creating anything on source pages
    ' (no origin marker), so per-page baselines cannot diverge.
    Dim sheetPages() As Page
    ReDim sheetPages(1 To totalSheets)

    Dim si As Integer
    For si = 1 To totalSheets
        ' Sheets are appended after the duplicated source pages.
        outDoc.AddPages 1
        Set outPage = outDoc.Pages(outDoc.Pages.Count)
        outPage.SetSize g_SheetWidth, g_SheetHeight

        ' Verify through the ACTUAL page API, not just the requested setting.
        Debug.Print "Output page sheet " & si & ": requested=" & _
                    Round(g_SheetWidth, 3) & " x " & Round(g_SheetHeight, 3) & _
                    " actual=" & Round(outPage.SizeWidth, 3) & " x " & Round(outPage.SizeHeight, 3) & _
                    " LeftX=" & Round(outPage.LeftX, 3) & " BottomY=" & Round(outPage.BottomY, 3) & _
                    " RightX=" & Round(outPage.RightX, 3) & " TopY=" & Round(outPage.TopY, 3)
        If Abs(outPage.SizeWidth - g_SheetWidth) > 0.05 Or _
           Abs(outPage.SizeHeight - g_SheetHeight) > 0.05 Then
            MsgBox "Output page size mismatch: requested " & _
                   Round(g_SheetWidth, 2) & " x " & Round(g_SheetHeight, 2) & _
                   " but page is " & Round(outPage.SizeWidth, 2) & " x " & _
                   Round(outPage.SizeHeight, 2) & ". Aborting.", vbExclamation
            Exit Function
        End If
        Set sheetPages(si) = outPage
    Next si

    slotIndex = LBound(orderArr)

    For sheetIndex = 1 To totalSheets
        Set outPage = sheetPages(sheetIndex)
        ' Re-read final frame (layout settled in Pass A).
        Debug.Print "Sheet " & sheetIndex & " final frame: LeftX=" & Round(outPage.LeftX, 3) & _
                    " BottomY=" & Round(outPage.BottomY, 3) & _
                    " RightX=" & Round(outPage.RightX, 3) & " TopY=" & Round(outPage.TopY, 3)

        For r = 0 To g_GridRows - 1
            For c = 0 To g_GridCols - 1

                If slotIndex <= UBound(orderArr) Then

                    Dim cellLeft As Double, cellBottom As Double
                    Dim pageNum As Integer
                    ' Absolute cell origin from the ACTUAL page frame.
                    ' X: left margin from LeftX. Y (upward): top margin down
                    ' from TopY, so Row 0 is at the top for any origin.
                    cellLeft = outPage.LeftX + g_MarginLeft + c * (cellW + g_GutterX)
                    cellBottom = outPage.TopY - g_MarginTop - cellH - r * (cellH + g_GutterY)

                    Debug.Print "Cell sheet=" & sheetIndex & " row=" & r & " col=" & c & _
                                " target X=" & Round(cellLeft, 3) & " Y(bottom)=" & Round(cellBottom, 3) & _
                                " W=" & Round(cellW, 3) & " H=" & Round(cellH, 3) & _
                                " right=" & Round(cellLeft + cellW, 3) & " top=" & Round(cellBottom + cellH, 3)

                    pageNum = orderArr(slotIndex)

                    If pageNum <> 0 And pageNum <= totalPages Then
                        PlacePageInCell srcDoc, pageNum, outDoc, outPage, _
                                        cellLeft, cellBottom, cellW, cellH, r, c
                    End If
                    ' pageNum = 0 means "blank slot" (e.g. padding page in a
                    ' signature) — cell is left empty but crop marks still
                    ' get drawn below so the sheet can still be trimmed.

                    If g_AddCropMarks Then
                        DrawCropMarks outDoc, outPage, cellLeft, cellBottom, cellW, cellH, g_BleedSize
                    End If

                    slotIndex = slotIndex + 1
                End If

            Next c
        Next r

    Next sheetIndex

    Application.Optimization = False
    Application.EventsEnabled = True

    ' Drop the duplicated source pages, leaving only the imposed sheets.
    Dim k As Integer
    On Error Resume Next
    For k = 1 To totalPages
        outDoc.Pages(1).Delete
    Next k
    Err.Clear
    On Error GoTo 0

    Application.Refresh

    MsgBox "Imposition complete (" & g_LayoutMode & "): " & totalSheets & _
           " sheet(s) created from " & totalPages & " source page(s).", vbInformation

    RunImposition = True

End Function

'--------------------------------------------------------------
' Checks trim+gutter+margin fit against the authoritative sheet.
' Bleed/crop length excluded by design (may extend into margins).
' Returns False + message box when it cannot fit at 100% (no scaling).
'--------------------------------------------------------------
Private Function ValidateSheetFits(cellW As Double, cellH As Double) As Boolean

    Dim reqW As Double, reqH As Double

    ValidateSheetFits = False

    ' Symmetric margin model (form has Margin L/T only):
    ' left==right==g_MarginLeft, top==bottom==g_MarginTop.
    reqW = 2 * g_MarginLeft + g_GridCols * cellW + (g_GridCols - 1) * g_GutterX
    reqH = 2 * g_MarginTop + g_GridRows * cellH + (g_GridRows - 1) * g_GutterY

    Debug.Print "ValidateSheetFits: required=" & Round(reqW, 3) & " x " & Round(reqH, 3) & _
                " available(sheet)=" & Round(g_SheetWidth, 3) & " x " & Round(g_SheetHeight, 3)

    If reqW > g_SheetWidth + 0.01 Or reqH > g_SheetHeight + 0.01 Then
        MsgBox "Layout does not fit the sheet at 100% (no scaling applied)." & vbCrLf & vbCrLf & _
               "Required (trim + gutter + margins): " & Round(reqW, 2) & " x " & Round(reqH, 2) & vbCrLf & _
               "Sheet: " & Round(g_SheetWidth, 2) & " x " & Round(g_SheetHeight, 2) & vbCrLf & vbCrLf & _
               "Cell: " & Round(cellW, 2) & " x " & Round(cellH, 2) & _
               "  Grid: " & g_GridRows & " x " & g_GridCols & vbCrLf & _
               "Gutter: " & Round(g_GutterX, 2) & " / " & Round(g_GutterY, 2) & _
               "  Margin L/T: " & Round(g_MarginLeft, 2) & " / " & Round(g_MarginTop, 2) & vbCrLf & vbCrLf & _
               "Increase sheet size, reduce grid/gutter/margins, or use a smaller trim.", _
               vbExclamation
        Exit Function
    End If

    ValidateSheetFits = True

End Function

'--------------------------------------------------------------
' Fills orderArr(1 To n) with the source-page index to place in
' each successive cell/slot, according to g_LayoutMode. A value
' of 0 means "leave this slot blank". Returns False (and shows a
' message box) if settings are invalid for the chosen mode.
'--------------------------------------------------------------
Private Function BuildPageOrder(totalPages As Integer, pagesPerSheet As Integer, _
                         ByRef orderArr() As Integer) As Boolean

    Dim i As Integer, s As Integer, idx As Integer

    Select Case g_LayoutMode

        Case "Simple"
            ReDim orderArr(1 To totalPages)
            For i = 1 To totalPages
                orderArr(i) = i
            Next i

        Case "StepRepeat"
            Dim totalSlots As Integer
            totalSlots = g_StepRepeatSheetCount * pagesPerSheet
            If g_StepRepeatPage < 1 Or g_StepRepeatPage > totalPages Then
                MsgBox "Step & Repeat page must be between 1 and " & totalPages & ".", vbExclamation
                BuildPageOrder = False
                Exit Function
            End If
            ReDim orderArr(1 To totalSlots)
            For i = 1 To totalSlots
                orderArr(i) = g_StepRepeatPage
            Next i

        Case "Signature"
            If pagesPerSheet <> 2 Then
                MsgBox "Signature mode currently supports a 2-up layout only " & _
                       "(set Grid Rows = 1, Grid Cols = 2). More complex " & _
                       "n-up signature schemes need a different fold/order formula.", _
                       vbExclamation
                BuildPageOrder = False
                Exit Function
            End If

            ' Pad total pages up to a multiple of 4 (blank pages are
            ' inserted at the end of the book, standard for saddle stitch).
            Dim paddedTotal As Integer
            paddedTotal = totalPages
            Do While paddedTotal Mod 4 <> 0
                paddedTotal = paddedTotal + 1
            Loop

            Dim numSheets As Integer
            numSheets = paddedTotal / 4

            ReDim orderArr(1 To paddedTotal)
            idx = 1
            For s = 0 To numSheets - 1
                Dim frontLeft As Integer, frontRight As Integer
                Dim backLeft As Integer, backRight As Integer
                frontLeft = paddedTotal - 2 * s
                frontRight = 2 * s + 1
                backLeft = 2 * s + 2
                backRight = paddedTotal - 2 * s - 1

                orderArr(idx) = IIf(frontLeft <= totalPages, frontLeft, 0): idx = idx + 1
                orderArr(idx) = IIf(frontRight <= totalPages, frontRight, 0): idx = idx + 1
                orderArr(idx) = IIf(backLeft <= totalPages, backLeft, 0): idx = idx + 1
                orderArr(idx) = IIf(backRight <= totalPages, backRight, 0): idx = idx + 1
            Next s

            ' Order produced is: Sheet1-Front(L,R), Sheet1-Back(L,R),
            ' Sheet2-Front(L,R), Sheet2-Back(L,R), ...
            ' i.e. each pair of output pages (front, back) = one physical
            ' sheet of paper, printed 2-up, then folded and nested/stapled.

        Case Else
            MsgBox "Unknown Layout Mode: " & g_LayoutMode & _
                   ". Use ""Simple"", ""StepRepeat"", or ""Signature"".", vbCritical
            BuildPageOrder = False
            Exit Function

    End Select

    BuildPageOrder = True

End Function

'--------------------------------------------------------------
' Copies all shapes from the source page into outPage at 100% (no
' scaling). NO origin marker is used: CopyToLayer preserves each
' source page's absolute document coordinates, so pages copied from
' different source pages land at different absolute baselines. The
' only reliable baseline is the ACTUAL post-paste bounding box, which
' is measured here and moved to the desired cell (preserving the
' source content-to-trim offset). All geometry uses absolute page
' edges (LeftX/RightX/TopY/BottomY); Move(dx,dy) is reference-point
' independent. No hard-coded offsets, no (0,0) assumption.
'--------------------------------------------------------------
Private Sub PlacePageInCell(srcDoc As Document, srcPageIdx As Integer, _
                     outDoc As Document, outPage As Page, _
                     cellLeft As Double, cellBottom As Double, _
                     cellW As Double, cellH As Double, _
                     cellRow As Integer, cellCol As Integer)

    Dim srcPage As Page
    Dim srcShapes As ShapeRange
    Dim originalActiveDoc As Document

    Set originalActiveDoc = Application.ActiveDocument
    Set srcPage = outDoc.Pages(srcPageIdx)

    Debug.Print "--- Place src=" & srcPageIdx & " -> row=" & cellRow & " col=" & cellCol & " ---"
    Debug.Print "Source page: LeftX=" & Round(srcPage.LeftX, 3) & _
                " RightX=" & Round(srcPage.RightX, 3) & _
                " TopY=" & Round(srcPage.TopY, 3) & _
                " BottomY=" & Round(srcPage.BottomY, 3) & _
                " Width=" & Round(srcPage.SizeWidth, 3) & _
                " Height=" & Round(srcPage.SizeHeight, 3)
    Debug.Print "Output page: LeftX=" & Round(outPage.LeftX, 3) & _
                " RightX=" & Round(outPage.RightX, 3) & _
                " TopY=" & Round(outPage.TopY, 3) & _
                " BottomY=" & Round(outPage.BottomY, 3) & _
                " Width=" & Round(outPage.SizeWidth, 3) & _
                " Height=" & Round(outPage.SizeHeight, 3)
    Debug.Print "Cell: row=" & cellRow & " col=" & cellCol & _
                " target X=" & Round(cellLeft, 3) & _
                " Y(bottom)=" & Round(cellBottom, 3) & _
                " W=" & Round(cellW, 3) & " H=" & Round(cellH, 3) & _
                " right=" & Round(cellLeft + cellW, 3) & _
                " top=" & Round(cellBottom + cellH, 3)

    If srcPage.Shapes.Count = 0 Then
        Debug.Print "Place src=" & srcPageIdx & ": source page has no shapes, cell left empty."
        Exit Sub
    End If

    ' Measure source content bbox BEFORE copy, to preserve the
    ' content-to-trim offset (content min minus page origin). For a
    ' full-bleed page this offset is ~0; for inset content it keeps the
    ' original layout instead of snapping content to the cell corner.
    Dim so As Shape
    Dim srcMinL As Double, srcMinB As Double, srcMaxR As Double, srcMaxT As Double
    Dim srcFirst As Boolean
    srcFirst = True
    For Each so In srcPage.Shapes
        If srcFirst Then
            srcMinL = so.LeftX
            srcMaxR = so.RightX
            srcMinB = so.BottomY
            srcMaxT = so.TopY
            srcFirst = False
        Else
            If so.LeftX < srcMinL Then srcMinL = so.LeftX
            If so.RightX > srcMaxR Then srcMaxR = so.RightX
            If so.BottomY < srcMinB Then srcMinB = so.BottomY
            If so.TopY > srcMaxT Then srcMaxT = so.TopY
        End If
    Next so
    Dim srcOffX As Double, srcOffY As Double
    srcOffX = srcMinL - srcPage.LeftX
    srcOffY = srcMinB - srcPage.BottomY
    Debug.Print "Source content bbox: LeftX=" & Round(srcMinL, 3) & _
                " RightX=" & Round(srcMaxR, 3) & _
                " TopY=" & Round(srcMaxT, 3) & _
                " BottomY=" & Round(srcMinB, 3) & _
                " Width=" & Round(srcMaxR - srcMinL, 3) & _
                " Height=" & Round(srcMaxT - srcMinB, 3) & _
                " offsetFromPageOrigin dx=" & Round(srcOffX, 3) & _
                " dy=" & Round(srcOffY, 3)

    Set srcShapes = srcPage.Shapes.All
    If srcShapes Is Nothing Then Exit Sub

    ' Record which shapes are ALREADY on the sheet page before this copy, by their
    ' stable StaticID. CorelDRAW inserts freshly copied shapes at the FRONT of
    ' the z-order (new Shapes(1)), so index-based beforeCount+1..Count would
    ' re-grab the PREVIOUS page's shapes on the second call.
    Dim beforeIDs As Collection
    Set beforeIDs = New Collection
    For Each so In outPage.Shapes
        On Error Resume Next
        beforeIDs.Add True, CStr(so.StaticID)
        On Error GoTo 0
    Next so

    ' Same-document copy onto the sheet page's layer. This is the only transfer
    ' that works in this object model, so the whole imposition runs inside the
    ' duplicate; source pages are deleted at the end.
    srcShapes.CopyToLayer outPage.ActiveLayer

    ' Collect whichever shapes on the sheet page are NOT in beforeIDs -- these
    ' are the ones this call just added, wherever CorelDRAW inserted them.
    Dim newShapes() As Shape
    Dim nNew As Long
    Dim idx As Long
    Dim wasOld As Boolean
    nNew = 0
    If outPage.Shapes.Count > 0 Then
        ReDim newShapes(1 To outPage.Shapes.Count)
        For Each so In outPage.Shapes
            wasOld = True
            On Error Resume Next
            Err.Clear
            beforeIDs.Item CStr(so.StaticID)
            wasOld = (Err.Number = 0)
            On Error GoTo 0
            If Not wasOld Then
                nNew = nNew + 1
                Set newShapes(nNew) = so
            End If
        Next so
    End If
    If nNew > 0 Then
        ReDim Preserve newShapes(1 To nNew)
    Else
        Erase newShapes
    End If

    If nNew = 0 Then
        Debug.Print "Place src=" & srcPageIdx & ": WARNING nothing pasted" & _
                    " (shapes on source: " & CStr(srcPage.Shapes.Count) & ")"
        Exit Sub
    End If

    ' ACTUAL post-paste bounding box (union of just-pasted shapes), BEFORE move.
    ' This is the only trustworthy baseline: it differs per source page when
    ' CopyToLayer preserves absolute document coordinates.
    Dim pstL As Double, pstB As Double, pstR As Double, pstT As Double
    pstL = newShapes(1).LeftX
    pstR = newShapes(1).RightX
    pstB = newShapes(1).BottomY
    pstT = newShapes(1).TopY
    For idx = 2 To nNew
        If newShapes(idx).LeftX < pstL Then pstL = newShapes(idx).LeftX
        If newShapes(idx).RightX > pstR Then pstR = newShapes(idx).RightX
        If newShapes(idx).BottomY < pstB Then pstB = newShapes(idx).BottomY
        If newShapes(idx).TopY > pstT Then pstT = newShapes(idx).TopY
    Next idx
    Debug.Print "Pasted (pre-move) bbox src=" & srcPageIdx & _
                ": LeftX=" & Round(pstL, 3) & _
                " RightX=" & Round(pstR, 3) & _
                " TopY=" & Round(pstT, 3) & _
                " BottomY=" & Round(pstB, 3) & _
                " Width=" & Round(pstR - pstL, 3) & _
                " Height=" & Round(pstT - pstB, 3)

    ' Desired union position: cell origin plus the preserved source offset,
    ' so the source page origin maps exactly to the cell origin at 100%.
    Dim wantL As Double, wantB As Double
    Dim correctionX As Double, correctionY As Double
    wantL = cellLeft + srcOffX
    wantB = cellBottom + srcOffY
    correctionX = wantL - pstL
    correctionY = wantB - pstB
    Debug.Print "Place src=" & srcPageIdx & _
                ": wantLeft=" & Round(wantL, 3) & _
                " wantBottom=" & Round(wantB, 3) & _
                " correction dx=" & Round(correctionX, 3) & _
                " dy=" & Round(correctionY, 3)

    For idx = 1 To nNew
        newShapes(idx).Move correctionX, correctionY
    Next idx

    ' Verify AFTER move: union of placed shapes must be inside the sheet.
    Dim pl As Double, pb As Double, pr As Double, pt As Double
    pl = newShapes(1).LeftX
    pr = newShapes(1).RightX
    pb = newShapes(1).BottomY
    pt = newShapes(1).TopY
    For idx = 2 To nNew
        If newShapes(idx).LeftX < pl Then pl = newShapes(idx).LeftX
        If newShapes(idx).RightX > pr Then pr = newShapes(idx).RightX
        If newShapes(idx).BottomY < pb Then pb = newShapes(idx).BottomY
        If newShapes(idx).TopY > pt Then pt = newShapes(idx).TopY
    Next idx
    Debug.Print "Placed page src=" & srcPageIdx & _
                ": actual LeftX=" & Round(pl, 3) & _
                " BottomY=" & Round(pb, 3) & _
                " RightX=" & Round(pr, 3) & _
                " TopY=" & Round(pt, 3)
    If pl < outPage.LeftX - 0.01 Or pr > outPage.RightX + 0.01 Or _
       pb < outPage.BottomY - 0.01 Or pt > outPage.TopY + 0.01 Then
        Debug.Print "Placed page src=" & srcPageIdx & ": WARNING content outside sheet frame!"
    Else
        Debug.Print "Placed page src=" & srcPageIdx & ": inside sheet OK."
    End If
    Debug.Print "Placed trim check src=" & srcPageIdx & _
                ": cell X=" & Round(cellLeft, 3) & _
                " Y=" & Round(cellBottom, 3) & _
                " cellRight=" & Round(cellLeft + cellW, 3) & _
                " cellTop=" & Round(cellBottom + cellH, 3)

    originalActiveDoc.Activate

End Sub

'--------------------------------------------------------------
' Draws 8 short crop-mark lines (2 per corner) just outside the
' cell's trim box, offset by bleedSize. Cell corner is absolute
' (cellLeft, cellBottom); each segment is re-measured after creation
' and nudged so marks stay glued to the cell even if the creation
' call interprets coordinates page-locally.
'--------------------------------------------------------------
Private Sub DrawCropMarks(outDoc As Document, outPage As Page, _
                   cellLeft As Double, cellBottom As Double, _
                   cellW As Double, cellH As Double, bleedSize As Double)

    Dim lyr As Layer

    outDoc.Activate
    outPage.Activate
    Set lyr = outPage.ActiveLayer

    Dim corners(3, 1) As Double ' 4 corners: (x, y) of each trim corner (absolute)
    corners(0, 0) = cellLeft:           corners(0, 1) = cellBottom               ' bottom-left
    corners(1, 0) = cellLeft + cellW:   corners(1, 1) = cellBottom               ' bottom-right
    corners(2, 0) = cellLeft:           corners(2, 1) = cellBottom + cellH       ' top-left
    corners(3, 0) = cellLeft + cellW:   corners(3, 1) = cellBottom + cellH       ' top-right

    Dim i As Integer
    For i = 0 To 3
        Dim cx As Double, cy As Double
        Dim signX As Integer, signY As Integer
        cx = corners(i, 0): cy = corners(i, 1)
        signX = IIf(i = 1 Or i = 3, 1, -1)
        signY = IIf(i = 2 Or i = 3, 1, -1)

        ' Intended absolute endpoints.
        Dim hx1 As Double, hy1 As Double, hx2 As Double, hy2 As Double
        Dim vx1 As Double, vy1 As Double, vx2 As Double, vy2 As Double
        hx1 = cx + signX * bleedSize
        hy1 = cy
        hx2 = cx + signX * (bleedSize + g_CropMarkLength)
        hy2 = cy
        vx1 = cx
        vy1 = cy + signY * bleedSize
        vx2 = cx
        vy2 = cy + signY * (bleedSize + g_CropMarkLength)

        Dim hLine As Shape, vLine As Shape
        Set hLine = lyr.CreateLineSegment(hx1, hy1, hx2, hy2)
        Set vLine = lyr.CreateLineSegment(vx1, vy1, vx2, vy2)

        ' Self-correct: if Create* interprets coords page-locally while we
        ' passed absolute (or vice versa), the segment will be off by exactly
        ' the page-origin offset. Re-measure via absolute edges and Move the
        ' difference so the mark lands on the intended absolute endpoints.
        ' Lengths are preserved by construction, only translation can differ.
        CorrectMarkPosition hLine, hx1, hy1, hx2, hy2
        CorrectMarkPosition vLine, vx1, vy1, vx2, vy2
    Next i

End Sub

'--------------------------------------------------------------
' Nudges a freshly created crop-mark segment so its absolute bbox
' matches the intended absolute endpoints. No-op when creation
' already used absolute coords.
'--------------------------------------------------------------
Private Sub CorrectMarkPosition(mark As Shape, _
                                sx As Double, sy As Double, _
                                ex As Double, ey As Double)
    On Error Resume Next
    If mark Is Nothing Then Exit Sub
    Dim wantLeft As Double, wantRight As Double
    Dim wantBottom As Double, wantTop As Double
    If sx < ex Then
        wantLeft = sx
        wantRight = ex
    Else
        wantLeft = ex
        wantRight = sx
    End If
    If sy < ey Then
        wantBottom = sy
        wantTop = ey
    Else
        wantBottom = ey
        wantTop = sy
    End If
    Dim dx As Double, dy As Double
    dx = wantLeft - mark.LeftX
    dy = wantBottom - mark.BottomY
    ' Ignore sub-micron rounding; correct only real origin offsets.
    If Abs(dx) > 0.001 Or Abs(dy) > 0.001 Then
        mark.Move dx, dy
    End If
    On Error GoTo 0
End Sub
