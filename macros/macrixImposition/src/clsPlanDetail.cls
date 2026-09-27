Attribute VB_Name = "clsPlanDetail"
'==============================================================
' clsPlanDetail
' One sheet type within a GangPlanResult. For "shared" plans this is
' the single sheet layout. For "mixed" plans each entry is a distinct
' sheet composition that is used a certain number of times.
'
' jobLayouts: Collection of clsJobLayout, one per job on this sheet
'             type. Always at least one.
'==============================================================

Option Explicit

Public sheetW As Double   ' usable width on this sheet type (media - 2*margin)
Public sheetH As Double   ' usable height on this sheet type
Public sheetsUsed As Long ' how many sheets of this type are in the plan

Public jobLayouts As Collection ' Collection of clsJobLayout

Public Sub Init()
    Set jobLayouts = New Collection
    sheetsUsed = 0
    sheetW = 0
    sheetH = 0
End Sub
