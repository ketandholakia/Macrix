'==============================================================
' CorelDRAW Imposition Macro (Simple N-up / Step & Repeat / Signature)
'
' HOW TO USE:
' 1. Open CorelDRAW, open the document you want to impose.
' 2. Tools > Macros > Visual Basic Editor (Alt+F11).
' 3. Insert > Module, paste this whole file in.
' 4. Edit the constants in "SETTINGS" below to match your job,
'    including LAYOUT_MODE ("Simple", "StepRepeat", or "Signature").
' 5. Run "RunImposition" (F5, or Tools > Macros > Run Macro).
' 6. A NEW document is created with the imposed sheets. Your
'    original document is left untouched.
'
' NOTE: CorelDRAW's VBA object model varies slightly between
' versions (X7, 2018, 2021, 2024...). If a line errors, it is
' almost always SetPosition / Layer assignment / CreateLineSegment
' — check the equivalent method name in your version's VBA object
' browser (F2 in the VBA IDE, search "Shape" or "Layer").
'==============================================================

Option Explicit

'---------------- SETTINGS ----------------
' All values are in the unit set by UNIT_TO_USE below.
Const UNIT_TO_USE As cdrUnit = cdrMillimeter

' LAYOUT_MODE options:
'   "Simple"     - straightforward N-up, sequential page order
'   "StepRepeat" - one source page repeated across every cell/sheet
'   "Signature"  - saddle-stitch booklet order (2-up only: GRID_ROWS=1, GRID_COLS=2)
Const LAYOUT_MODE As String = "Simple"

Const SHEET_WIDTH As Double = 320      ' target output sheet width
Const SHEET_HEIGHT As Double = 450     ' target output sheet height

Const GRID_ROWS As Integer = 2         ' pages per column (vertical count)
Const GRID_COLS As Integer = 2         ' pages per row (horizontal count)

Const GUTTER_X As Double = 5           ' horizontal gap between pages
Const GUTTER_Y As Double = 5           ' vertical gap between pages

Const MARGIN_LEFT As Double = 10       ' left margin of sheet
Const MARGIN_TOP As Double = 10        ' top margin of sheet

Const BLEED_SIZE As Double = 3         ' bleed for crop-mark offset

Const ADD_CROP_MARKS As Boolean = True
Const CROP_MARK_LENGTH As Double = 5   ' length of each crop mark line

' --- Only used when LAYOUT_MODE = "StepRepeat" ---
Const STEP_REPEAT_PAGE As Integer = 1     ' which source page to repeat everywhere
Const STEP_REPEAT_SHEET_COUNT As Integer = 4  ' how many output sheets to produce
'-------------------------------------------

Sub RunImposition()

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

    If ActiveDocument Is Nothing Then
        MsgBox "Open a document first.", vbExclamation
        Exit Sub
    End If

    Set srcDoc = ActiveDocument
    srcDoc.Unit = UNIT_TO_USE
    totalPages = srcDoc.Pages.Count

    If totalPages = 0 Then
        MsgBox "The active document has no pages.", vbExclamation
        Exit Sub
    End If

    pagesPerSheet = GRID_ROWS * GRID_COLS

    ' Build the sequence of source-page indices to drop into each
    ' cell, slot by slot (0 = leave this slot blank).
    If Not BuildPageOrder(totalPages, pagesPerSheet, orderArr) Then
        Exit Sub ' BuildPageOrder already showed an error message
    End If
    orderLen = UBound(orderArr) - LBound(orderArr) + 1
    totalSheets = Int((orderLen - 1) / pagesPerSheet) + 1

    ' Cell size: divide usable sheet area evenly, accounting for gutters
    cellW = (SHEET_WIDTH - 2 * MARGIN_LEFT - (GRID_COLS - 1) * GUTTER_X) / GRID_COLS
    cellH = (SHEET_HEIGHT - 2 * MARGIN_TOP - (GRID_ROWS - 1) * GUTTER_Y) / GRID_ROWS

    Application.Optimization = True
    Application.EventsEnabled = False

    ' Create a fresh output document
    Set outDoc = CreateDocument
    outDoc.Unit = UNIT_TO_USE
    outDoc.ReferencePoint = cdrTopLeft

    slotIndex = LBound(orderArr)

    For sheetIndex = 1 To totalSheets

        ' First sheet reuses the doc's default page; later ones are added
        If sheetIndex = 1 Then
            Set outPage = outDoc.Pages(1)
        Else
            Set outPage = outDoc.Pages.Add(outDoc.Pages.Count)
        End If
        outPage.SetSize SHEET_WIDTH, SHEET_HEIGHT

        For r = 0 To GRID_ROWS - 1
            For c = 0 To GRID_COLS - 1

                If slotIndex <= UBound(orderArr) Then

                    Dim cellX As Double, cellY As Double
                    Dim pageNum As Integer
                    cellX = MARGIN_LEFT + c * (cellW + GUTTER_X)
                    cellY = MARGIN_TOP + r * (cellH + GUTTER_Y)
                    pageNum = orderArr(slotIndex)

                    If pageNum <> 0 And pageNum <= totalPages Then
                        PlacePageInCell srcDoc, pageNum, outDoc, outPage, _
                                        cellX, cellY, cellW, cellH
                    End If
                    ' pageNum = 0 means "blank slot" (e.g. padding page in a
                    ' signature) — cell is left empty but crop marks still
                    ' get drawn below so the sheet can still be trimmed.

                    If ADD_CROP_MARKS Then
                        DrawCropMarks outDoc, outPage, cellX, cellY, cellW, cellH, BLEED_SIZE
                    End If

                    slotIndex = slotIndex + 1
                End If

            Next c
        Next r

    Next sheetIndex

    Application.Optimization = False
    Application.EventsEnabled = True
    Application.Refresh

    MsgBox "Imposition complete (" & LAYOUT_MODE & "): " & totalSheets & _
           " sheet(s) created from " & totalPages & " source page(s).", vbInformation

End Sub

'--------------------------------------------------------------
' Fills orderArr(1 To n) with the source-page index to place in
' each successive cell/slot, according to LAYOUT_MODE. A value
' of 0 means "leave this slot blank". Returns False (and shows a
' message box) if settings are invalid for the chosen mode.
'--------------------------------------------------------------
Function BuildPageOrder(totalPages As Integer, pagesPerSheet As Integer, _
                         ByRef orderArr() As Integer) As Boolean

    Dim i As Integer, s As Integer, idx As Integer

    Select Case LAYOUT_MODE

        Case "Simple"
            ReDim orderArr(1 To totalPages)
            For i = 1 To totalPages
                orderArr(i) = i
            Next i

        Case "StepRepeat"
            Dim totalSlots As Integer
            totalSlots = STEP_REPEAT_SHEET_COUNT * pagesPerSheet
            If STEP_REPEAT_PAGE < 1 Or STEP_REPEAT_PAGE > totalPages Then
                MsgBox "STEP_REPEAT_PAGE must be between 1 and " & totalPages & ".", vbExclamation
                BuildPageOrder = False
                Exit Function
            End If
            ReDim orderArr(1 To totalSlots)
            For i = 1 To totalSlots
                orderArr(i) = STEP_REPEAT_PAGE
            Next i

        Case "Signature"
            If pagesPerSheet <> 2 Then
                MsgBox "Signature mode currently supports a 2-up layout only " & _
                       "(set GRID_ROWS = 1, GRID_COLS = 2). More complex " & _
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
            MsgBox "Unknown LAYOUT_MODE: " & LAYOUT_MODE & _
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
Sub PlacePageInCell(srcDoc As Document, srcPageIdx As Integer, _
                     outDoc As Document, outPage As Page, _
                     cellX As Double, cellY As Double, _
                     cellW As Double, cellH As Double)

    Dim srcPage As Page
    Dim srcShapes As ShapeRange
    Dim dupShapes As ShapeRange
    Dim originalActiveDoc As Document

    Set originalActiveDoc = Application.ActiveDocument
    Set srcPage = srcDoc.Pages(srcPageIdx)

    If srcPage.Shapes.Count = 0 Then Exit Sub ' nothing to place

    ' Switch context to source doc to select + copy its page content
    srcDoc.Activate
    Set srcShapes = srcPage.Shapes.All
    srcShapes.Copy

    ' Switch to output doc, paste onto the target page
    outDoc.Activate
    outPage.Activate
    Set dupShapes = outDoc.Paste

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

'--------------------------------------------------------------
' Draws 8 short crop-mark lines (2 per corner) just outside the
' cell's trim box, offset by bleedSize.
'--------------------------------------------------------------
Sub DrawCropMarks(outDoc As Document, outPage As Page, _
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
                               cx + signX * (bleedSize + CROP_MARK_LENGTH), cy
        ' vertical mark
        lyr.CreateLineSegment cx, cy + signY * bleedSize, _
                               cx, cy + signY * (bleedSize + CROP_MARK_LENGTH)
    Next i

End Sub
