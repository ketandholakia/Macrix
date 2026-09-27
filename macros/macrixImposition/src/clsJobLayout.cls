Attribute VB_Name = "clsJobLayout"
'==============================================================
' clsJobLayout
' How one clsGangJob is laid out on ONE sheet type (one "detail" in a
' plan). Expresses the grid geometry and piece counts as data only --
' no CorelDRAW objects.
'==============================================================

Option Explicit

Public jobID As Long
Public jobName As String

Public cols As Long      ' columns on this sheet for this job
Public rows As Long      ' rows on this sheet for this job
Public piecesPerSheet As Long  ' cols * rows

Public rotated As Boolean   ' True if the job is rotated 90 deg on this sheet
Public effW As Double       ' effective width used for grid (rotated? Height : Width)
Public effH As Double       ' effective height used for grid

Public shareW As Double   ' width of the slice this job occupies on the sheet (for shared plans)
Public shareH As Double   ' height of the slice this job occupies on the sheet

Public Sub Init()
    cols = 0
    rows = 0
    piecesPerSheet = 0
    rotated = False
    effW = 0
    effH = 0
    shareW = 0
    shareH = 0
    jobID = 0
    jobName = ""
End Sub
