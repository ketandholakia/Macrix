Attribute VB_Name = "modMain"
Option Explicit
''
' modMain -- the single public entry point for this macro.
'
' Only this module is listed in CorelDRAW's macro list. Everything else lives in
' Option Private Module modules, which stay callable everywhere inside the project
' (the form included) but do not clutter the macro list.
''
' Depends on: modSettings (InitDefaultSettings), modImposition, frmImpositionSettings.
'             modTileFillSettings (InitDefaultTileSettings), modTileFill,
'             frmTileFillSettings.
'             modGangJobSettings (InitDefaultGangSettings), modGangJob, clsGangJob
'             (Phase 1 -- data model only, no form yet).

Public Sub ShowImpositionForm()
    ' Load once per session: InitDefaultSettings resets to built-in defaults
    ' and overlays last-used registry values. Guarded so reopening the form
    ' never wipes the in-memory last-used values.
    If Not g_SettingsInitialized Then InitDefaultSettings
    ' frmImpositionSettings is a real design-time form in this project (built by
    ' scripts\Build-Form.ps1), so it is shown directly: no runtime form generation
    ' and no VBE access required. Its "Run" button calls modImposition.RunImposition.
    frmImpositionSettings.Show
End Sub

Public Sub ShowTileFillForm()
    ' Same pattern as ShowImpositionForm: InitDefaultTileSettings resets to
    ' built-in defaults and overlays last-used registry values, guarded so
    ' reopening the form never wipes in-memory last-used values.
    If Not g_TileSettingsInitialized Then InitDefaultTileSettings
    ' frmTileFillSettings is built the same way frmImpositionSettings is (see
    ' mdlTileFillFormBuilder.BuildTileFillFormWithCode) -- run that once from
    ' the Immediate Window (or via Build-Form.ps1, if that script is extended
    ' to target it) before this will find the form. Its "Tile Fill Selection"
    ' button calls modTileFill.RunTileFill.
    frmTileFillSettings.Show
End Sub

' --- Gang Job (Phase 1: data model only -- no form yet) ---------------
'
' Select artwork in CorelDRAW, run GangJob_AddFromSelection, repeat per
' job, then run GangJob_ShowJobs (Ctrl+G for the Immediate Window) to
' confirm the list. These three will be replaced by a single
' ShowGangJobForm once frmGangJobSettings exists (Phase 6), the same
' way ShowTileFillForm now stands in place of raw InputBox prompts.

Public Sub GangJob_AddFromSelection()
    modGangJob.AddJobFromSelection
End Sub

Public Sub GangJob_ShowJobs()
    modGangJob.DumpGangJobs
End Sub

Public Sub GangJob_ClearJobs()
    modGangJob.ClearGangJobs
End Sub

' Rebuilds the job list purely from artwork already tagged (via a
' previous GangJob_AddFromSelection) in the current document -- no
' re-typing of Name/Quantity/Rotation needed. Run GangJob_ShowJobs
' afterward to confirm what it found.
Public Sub GangJob_ScanDocument()
    modGangJob.ScanDocumentForGangJobs
End Sub

' --- Gang Job (Phase 2+3: optimizer) -----------------------------------
'
' GangJob_Optimize runs the pure-data optimizer (modGangJobOptimizer)
' against the current in-memory g_GangJobs collection and dumps the
' best plan to the Immediate Window (Ctrl+G). No CorelDRAW page
' creation -- that is Phase 6 (modGangJobOutput).
'
' GangJob_OptimizeDump dumps ALL candidate plans (not just the best),
' useful for understanding the trade-offs the optimizer considered.

Public Sub GangJob_Optimize()
    If Not g_GangSettingsInitialized Then InitDefaultGangSettings
    If g_GangJobs Is Nothing Or g_GangJobs.Count = 0 Then
        Debug.Print "GangJob_Optimize: no jobs to optimize."
        Exit Sub
    End If
    Dim plans As Collection
    Set plans = modGangJobOptimizer.OptimizeGangJobs
    If plans.Count = 0 Then
        Debug.Print "GangJob_Optimize: optimizer returned no plans."
        Exit Sub
    End If
    Debug.Print "GangJob_Optimize: best plan (waste " & _
                    Format(plans(1).-wastePercent, "0.00%") & "):"
    modGangJobOptimizer.DumpGangPlan plans(1)
End Sub

Public Sub GangJob_OptimizeDump()
    If Not g_GangSettingsInitialized Then InitDefaultGangSettings
    If g_GangJobs Is Nothing Or g_GangJobs.Count = 0 Then
        Debug.Print "GangJob_OptimizeDump: no jobs to optimize."
        Exit Sub
    End If
    Dim plans As Collection
    Set plans = modGangJobOptimizer.OptimizeGangJobs
    modGangJobOptimizer.DumpAllGangPlans plans
End Sub
