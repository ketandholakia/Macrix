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
'==============================================================

Option Explicit
Option Private Module

Public Sub RunImposition()

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

    If Not g_SettingsInitialized Then InitDefaultSettings

    If ActiveDocument Is Nothing Then
        MsgBox "Open a document first.", vbExclamation
        Exit Sub
    End If

    Set srcDoc = ActiveDocument
    srcDoc.Unit = g_Unit
    totalPages = srcDoc.Pages.Count

    If totalPages = 0 Then
        MsgBox "The active document has no pages.", vbExclamation
        Exit Sub
    End If

    pagesPerSheet = g_GridRows * g_GridCols

    ' Build the sequence of source-page indices to drop into each
    ' cell, slot by slot (0 = leave this slot blank).
    If Not BuildPageOrder(totalPages, pagesPerSheet, orderArr) Then
        Exit Sub ' BuildPageOrder already showed an error message
    End If
    orderLen = UBound(orderArr) - LBound(orderArr) + 1
    totalSheets = Int((orderLen - 1) / pagesPerSheet) + 1

    ' Cell size: divide usable sheet area evenly, accounting for gutters
    cellW = (g_SheetWidth - 2 * g_MarginLeft - (g_GridCols - 1) * g_GutterX) / g_GridCols
    cellH = (g_SheetHeight - 2 * g_MarginTop - (g_GridRows - 1) * g_GutterY) / g_GridRows

    Application.Optimization = True
    Application.EventsEnabled = False

    ' Create a fresh output document
    Set outDoc = CreateDocument
    outDoc.Unit = g_Unit
    outDoc.ReferencePoint = cdrTopLeft

    ' The placement below copies and pastes shapes. That round-trip needs a fully live
    ' document context: with Optimization or EventsEnabled switched off, Copy silently
    ' copies nothing, so every Paste comes back empty and the sheet ends up holding only
    ' the crop marks. Diagnostics showed the source page had a shape while the sheet
    ' page count never increased, so both are switched back on for the loop.
    Application.Optimization = False
    Application.EventsEnabled = True

    slotIndex = LBound(orderArr)

    For sheetIndex = 1 To totalSheets

        ' First sheet reuses the doc's default page; later ones are added
        If sheetIndex = 1 Then
            Set outPage = outDoc.Pages(1)
        Else
            ' AddPages, not Pages.Add: the Pages collection has no Add method (error 438).
            outDoc.AddPages 1
            Set outPage = outDoc.Pages(outDoc.Pages.Count)
        End If
        outPage.SetSize g_SheetWidth, g_SheetHeight

        For r = 0 To g_GridRows - 1
            For c = 0 To g_GridCols - 1

                If slotIndex <= UBound(orderArr) Then

                    Dim cellX As Double, cellY As Double
                    Dim pageNum As Integer
                    cellX = g_MarginLeft + c * (cellW + g_GutterX)
                    cellY = g_MarginTop + r * (cellH + g_GutterY)
                    pageNum = orderArr(slotIndex)

                    If pageNum <> 0 And pageNum <= totalPages Then
                        PlacePageInCell srcDoc, pageNum, outDoc, outPage, _
                                        cellX, cellY, cellW, cellH
                    End If
                    ' pageNum = 0 means "blank slot" (e.g. padding page in a
                    ' signature) — cell is left empty but crop marks still
                    ' get drawn below so the sheet can still be trimmed.

                    If g_AddCropMarks Then
                        DrawCropMarks outDoc, outPage, cellX, cellY, cellW, cellH, g_BleedSize
                    End If

                    slotIndex = slotIndex + 1
                End If

            Next c
        Next r

    Next sheetIndex

    Application.Optimization = False
    Application.EventsEnabled = True
    Application.Refresh

    MsgBox "Imposition complete (" & g_LayoutMode & "): " & totalSheets & _
           " sheet(s) created from " & totalPages & " source page(s).", vbInformation

End Sub

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
' Duplicates all shapes from srcDoc's page (by index) into
' outPage at the given cell position (top-left of cell), without
' scaling (imposition normally preserves original page size).
'--------------------------------------------------------------
Private Sub PlacePageInCell(srcDoc As Document, srcPageIdx As Integer, _
                     outDoc As Document, outPage As Page, _
                     cellX As Double, cellY As Double, _
                     cellW As Double, cellH As Double)

    Dim srcPage As Page
    Dim srcShapes As ShapeRange
    Dim dupShapes As ShapeRange
    Dim originalActiveDoc As Document
    Dim beforeShapes As ShapeRange
    Dim beforeCount As Long
    Dim i As Long

    Set originalActiveDoc = Application.ActiveDocument
    Set srcPage = srcDoc.Pages(srcPageIdx)

    If srcPage.Shapes.Count = 0 Then Exit Sub ' nothing to place

    ' Copy the source page's shapes STRAIGHT onto the sheet page's layer.
    ' The clipboard route (Shapes.All.Copy + Layer.Paste) put nothing on the clipboard in
    ' this project, which is why the sheet page ended up holding only crop marks.
    Set srcShapes = srcPage.Shapes.All
    If srcShapes Is Nothing Then Exit Sub

    Set beforeShapes = outPage.Shapes.All
    beforeCount = beforeShapes.Count
    srcShapes.CopyToLayer outPage.ActiveLayer
    Set dupShapes = p_NewShapesSince(outPage, beforeCount)

    ' Fallback: clipboard round-trip, with the source shapes selected first.
    If dupShapes.Count = 0 Then
        srcDoc.Activate
        srcShapes.CreateSelection
        srcShapes.Copy
        outDoc.Activate
        outPage.Activate
        outPage.ActiveLayer.Paste
        Set dupShapes = p_NewShapesSince(outPage, beforeCount)
    End If

    If dupShapes.Count = 0 Then
        ' Nothing landed on the sheet page. Report the counts rather than failing
        ' silently (and do NOT fall back to Document.Selection -- it is a pointer-typed
        ' member that raises "Run-time error 13: Type mismatch" when assigned in VBA).
        Dim srcLocked As Boolean
        Dim srcLayer As String
        srcLocked = False
        srcLayer = "?"
        On Error Resume Next
        srcLocked = srcPage.Shapes(1).Locked
        srcLayer = srcPage.ActiveLayer.Name
        On Error GoTo 0
        MsgBox "Nothing was pasted for source page " & CStr(srcPageIdx) & "." & vbCrLf & vbCrLf & _
               "Shapes on the source page      : " & CStr(srcPage.Shapes.Count) & vbCrLf & _
               "First source shape locked      : " & CStr(srcLocked) & vbCrLf & _
               "Source active layer            : " & srcLayer & vbCrLf & _
               "Shapes on the sheet page before: " & CStr(beforeCount) & vbCrLf & _
               "Shapes on the sheet page after : " & CStr(outPage.Shapes.Count), _
               vbExclamation, "Imposition"
        Exit Sub
    End If

    ' Position: align pasted content's top-left to the cell's top-left.
    ' (Assumes outDoc.ReferencePoint = cdrTopLeft, set in RunImposition.)
    dupShapes.SetPosition cellX, cellY

    ' If the source page is bigger than the cell, you may want to scale
    ' it down to fit instead of clipping. Uncomment to enable fit-scaling:
    '
    ' Dim scaleFactor As Double
    ' scaleFactor = Application.Min(cellW / dupShapes.SizeWidth, cellH / dupShapes.SizeHeight)
    ' dupShapes.SetSize dupShapes.SizeWidth * scaleFactor, dupShapes.SizeHeight * scaleFactor
    ' dupShapes.SetPosition cellX, cellY

    originalActiveDoc.Activate

End Sub

' Collect the shapes added to a page since a given count (CopyToLayer/Paste return
' pointer-typed values that VBA can surface as Nothing, so count instead).
Private Function p_NewShapesSince(thePage As Page, ByVal sinceCount As Long) As ShapeRange
    On Error Resume Next
    Dim result As ShapeRange
    Dim k As Long
    Set result = New ShapeRange
    For k = sinceCount + 1 To thePage.Shapes.Count
        result.Add thePage.Shapes(k)
    Next k
    Set p_NewShapesSince = result
End Function

'--------------------------------------------------------------
' Draws 8 short crop-mark lines (2 per corner) just outside the
' cell's trim box, offset by bleedSize.
'--------------------------------------------------------------
Private Sub DrawCropMarks(outDoc As Document, outPage As Page, _
                   cellX As Double, cellY As Double, _
                   cellW As Double, cellH As Double, bleedSize As Double)

    Dim lyr As Layer

    outDoc.Activate
    outPage.Activate
    Set lyr = outPage.ActiveLayer

    Dim corners(3, 1) As Double ' 4 corners: (x, y) of each trim corner
    corners(0, 0) = cellX:           corners(0, 1) = cellY               ' top-left
    corners(1, 0) = cellX + cellW:   corners(1, 1) = cellY               ' top-right
    corners(2, 0) = cellX:           corners(2, 1) = cellY + cellH       ' bottom-left
    corners(3, 0) = cellX + cellW:   corners(3, 1) = cellY + cellH       ' bottom-right

    Dim i As Integer
    For i = 0 To 3
        Dim cx As Double, cy As Double
        Dim signX As Integer, signY As Integer
        cx = corners(i, 0): cy = corners(i, 1)
        signX = IIf(i = 1 Or i = 3, 1, -1)
        signY = IIf(i = 2 Or i = 3, 1, -1)

        ' horizontal mark
        lyr.CreateLineSegment cx + signX * bleedSize, cy, _
                               cx + signX * (bleedSize + g_CropMarkLength), cy
        ' vertical mark
        lyr.CreateLineSegment cx, cy + signY * bleedSize, _
                               cx, cy + signY * (bleedSize + g_CropMarkLength)
    Next i

End Sub
