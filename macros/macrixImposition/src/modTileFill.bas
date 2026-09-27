Attribute VB_Name = "modTileFill"
'==============================================================
' modTileFill
' Fills the active page with repeated copies of the CURRENT
' SELECTION, maximizing how many copies fit, without altering the
' original selected object(s) in any way (position, size, or
' grouping). The selection may contain multiple shapes (e.g. a
' logo made of several objects) -- ShapeRange operations treat the
' whole selection as one unit for sizing and duplicating, so
' nothing needs to be permanently grouped.
'
' Run: RunTileFill
' Optional settings tweak: PromptTileSettings (in
' modTileFillSettings.bas) -- margin, gutter, allow-rotate.
'
' DESIGN NOTE: modImposition's CopyToLayer (copying shapes from one
' PAGE to another) turned out to have an unpredictable paste
' baseline -- see that module's history. This macro avoids that
' failure mode entirely by never leaving the active page: it uses
' ShapeRange.Duplicate(), which hands back a direct reference to
' the new shapes (no index-guessing, no z-order ambiguity), and
' positions each duplicate via CenterX/CenterY, which -- unlike
' LeftX/TopY/PositionX/PositionY -- is not reference-point-
' dependent, so there's no top-left-vs-bottom-left ambiguity to
' get wrong.
'==============================================================

Option Explicit
Option Private Module

Public Function RunTileFill() As Boolean

    Dim doc As Document
    Dim pg As Page
    Dim sel As ShapeRange
    Dim selCount As Long

    Dim objW As Double, objH As Double
    Dim pageW As Double, pageH As Double
    Dim availW As Double, availH As Double

    Dim colsNormal As Long, rowsNormal As Long, countNormal As Long
    Dim colsRotated As Long, rowsRotated As Long, countRotated As Long
    Dim useRotated As Boolean

    Dim effW As Double, effH As Double
    Dim cols As Long, rows As Long
    Dim gridW As Double, gridH As Double
    Dim leftMargin As Double, bottomMargin As Double
    Dim r As Long, c As Long
    Dim placed As Long

    If Not g_TileSettingsInitialized Then InitDefaultTileSettings

    If ActiveDocument Is Nothing Then
        MsgBox "Open a document first.", vbExclamation
        Exit Function
    End If

    Set doc = ActiveDocument
    doc.Unit = g_TileUnit

    ' Guard against calling .Count on a Nothing SelectionRange -- VBA's "Or"
    ' does not short-circuit, so this has to be checked in two steps.
    selCount = 0
    On Error Resume Next
    selCount = doc.SelectionRange.Count
    On Error GoTo 0
    If selCount = 0 Then
        MsgBox "Select the object (or group of objects) to tile first.", vbExclamation
        Exit Function
    End If

    ' The selection, AS A WHOLE, is the unit to repeat. SizeWidth/SizeHeight
    ' give its combined bounding box; Duplicate()/CenterX/CenterY act on all
    ' member shapes together, preserving their relative arrangement -- so a
    ' multi-shape selection does not need to be Grouped to be treated as one
    ' tile, and the original is never touched.
    Set sel = doc.SelectionRange

    objW = sel.SizeWidth
    objH = sel.SizeHeight
    If objW <= 0 Or objH <= 0 Then
        MsgBox "Could not read the selection's size.", vbExclamation
        Exit Function
    End If

    Set pg = doc.ActivePage
    pageW = pg.SizeWidth
    pageH = pg.SizeHeight

    availW = pageW - 2 * g_TileMarginX
    availH = pageH - 2 * g_TileMarginY
    If availW <= 0 Or availH <= 0 Then
        MsgBox "The margin settings leave no usable space on this page.", vbExclamation
        Exit Function
    End If

    ' How many copies fit in the object's current orientation...
    colsNormal = FitCount(availW, objW, g_TileGutterX)
    rowsNormal = FitCount(availH, objH, g_TileGutterY)
    countNormal = colsNormal * rowsNormal

    ' ...versus rotated 90 degrees (width/height swapped for the fit test).
    colsRotated = FitCount(availW, objH, g_TileGutterX)
    rowsRotated = FitCount(availH, objW, g_TileGutterY)
    countRotated = colsRotated * rowsRotated

    useRotated = False
    If g_TileAllowRotate And countRotated > countNormal Then
        useRotated = True
    End If

    If useRotated Then
        cols = colsRotated: rows = rowsRotated
        effW = objH: effH = objW
    Else
        cols = colsNormal: rows = rowsNormal
        effW = objW: effH = objH
    End If

    If cols <= 0 Or rows <= 0 Then
        MsgBox "The selected object doesn't fit on the page even once " & _
               "with the current margin/gutter settings. Use PromptTileSettings " & _
               "to reduce them.", vbExclamation
        Exit Function
    End If

    ' Center the whole grid on the page -- any leftover space (the page size
    ' not being an exact multiple of the tile size) is split evenly on both
    ' sides rather than pushed to one edge.
    gridW = cols * effW + (cols - 1) * g_TileGutterX
    gridH = rows * effH + (rows - 1) * g_TileGutterY
    leftMargin = (pageW - gridW) / 2
    bottomMargin = (pageH - gridH) / 2

    Application.Optimization = True
    Application.EventsEnabled = False

    placed = 0
    Dim firstDup As ShapeRange
    Dim firstTargetX As Double, firstTargetY As Double

    For r = 0 To rows - 1
        For c = 0 To cols - 1

            Dim targetCenterX As Double, targetCenterY As Double
            targetCenterX = leftMargin + c * (effW + g_TileGutterX) + effW / 2
            targetCenterY = bottomMargin + r * (effH + g_TileGutterY) + effH / 2

            Dim dup As ShapeRange
            Set dup = sel.Duplicate()
            If useRotated Then dup.Rotate 90
            dup.CenterX = targetCenterX
            dup.CenterY = targetCenterY

            If placed = 0 Then
                Set firstDup = dup
                firstTargetX = targetCenterX
                firstTargetY = targetCenterY
            End If

            placed = placed + 1

        Next c
    Next r

    Application.Optimization = False
    Application.EventsEnabled = True
    Application.Refresh

    ' One-line sanity check on the very first duplicate -- if CenterX/CenterY
    ' ever turned out to be reference-point-sensitive after all (unlikely, but
    ' the imposition macro's history is a reminder not to assume), this will
    ' show a mismatch immediately instead of only showing up as a bad-looking
    ' page. Check View > Immediate Window (Ctrl+G) in the VBE.
    On Error Resume Next
    Debug.Print "RunTileFill: cols=" & cols & " rows=" & rows & _
                " rotated=" & useRotated & _
                "  first dup target=(" & Round(firstTargetX, 3) & "," & Round(firstTargetY, 3) & ")" & _
                "  actual=(" & Round(firstDup.CenterX, 3) & "," & Round(firstDup.CenterY, 3) & ")"
    On Error GoTo 0

    MsgBox "Tiled " & CStr(placed) & " cop" & IIf(placed = 1, "y", "ies") & _
           " (" & cols & " x " & rows & ")" & _
           IIf(useRotated, ", rotated 90 degrees to fit more.", ".") & vbCrLf & _
           "The original selection was left unchanged.", vbInformation

    RunTileFill = True

End Function

'--------------------------------------------------------------
' How many items of itemSize fit into available space, with
' gutter-sized gaps between them (n-1 gutters for n items).
'--------------------------------------------------------------
Private Function FitCount(available As Double, itemSize As Double, gutter As Double) As Long
    If itemSize <= 0 Then
        FitCount = 0
        Exit Function
    End If
    ' n*itemSize + (n-1)*gutter <= available  =>  n <= (available+gutter)/(itemSize+gutter)
    FitCount = Int((available + gutter) / (itemSize + gutter) + 0.0000001)
    If FitCount < 0 Then FitCount = 0
End Function
