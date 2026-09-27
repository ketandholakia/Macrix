Attribute VB_Name = "clsGangPlanResult"
'==============================================================
' clsGangPlanResult
' One candidate gang plan returned by modGangJobOptimizer.Phase 3.
' A plan describes ONE way to lay out ALL jobs onto sheets.
'
' strategy: "shared" = all jobs co-located on every sheet (one sheet
'           composition, sheetsNeeded = max of per-job sheet needs);
'           "mixed"  = different sheet types can be combined, each
'           sheet type is a "row" in planDetails.
'
' planDetails: Collection of clsPlanDetail, one per sheet type in the
'              plan. For "shared" there is exactly one detail. For
'              "mixed" there may be several.
'==============================================================

Option Explicit

Public strategy As String       ' "shared" or "mixed"
Public totalSheets As Long
Public totalWasteArea As Double
Public totalMediaArea As Double
Public wastePercent As Double

Public planDetails As Collection ' Collection of clsPlanDetail

Public Sub Init()
    Set planDetails = New Collection
    totalSheets = 0
    totalWasteArea = 0
    totalMediaArea = 0
    wastePercent = 0
    strategy = ""
End Sub
