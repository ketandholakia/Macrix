Attribute VB_Name = "modGangJobSettings"
'==============================================================
' modGangJobSettings
' Settings for the Gang Job data model (modGangJob) and optimizer
' (modGangJobOptimizer). Kept as its own registry section -- per
' docs/MACROS.md's note not to let macros' "last used" values bleed
' into each other.
'
' Now includes media size, margins, and gutters needed by the
' optimizer (Phase 2+), the same way modTileFillSettings grew
' alongside modTileFill.
'==============================================================

Option Explicit
Option Private Module

Private Const APP_NAME As String = "macrixGangJob"
Private Const SECTION_NAME As String = "Settings"

Public g_GangSettingsInitialized As Boolean

Public g_GangUnit As cdrUnit

' Auto-incrementing ID handed to each new clsGangJob, so jobs stay
' identifiable by RemoveGangJob even after earlier Add/Remove churn
' (unlike a plain Collection index, which shifts).
Public g_GangNextJobID As Long

' --------------------------------------------------------------
' Media size for the optimizer (Phase 2+). These define the sheet
' size the optimizer thinks it is working with. For now these are
' fixed defaults; a future frmGangJobSettings form will let the user
' override them.
' --------------------------------------------------------------
Public g_GangMediaWidth As Double
Public g_GangMediaHeight As Double

' Margins: gap kept between the grid edge and the media edge.
Public g_GangMarginX As Double
Public g_GangMarginY As Double

' Gutter: gap kept between adjacent copies in the grid.
Public g_GangGutterX As Double
Public g_GangGutterY As Double

' --------------------------------------------------------------
' Resets to built-in defaults, then overlays last-used values from
' the registry if any were saved by a previous session. Guarded
' (like InitDefaultSettings / InitDefaultTileSettings) so calling
' this again mid-session never wipes in-memory values.
' --------------------------------------------------------------
Sub InitDefaultGangSettings()

    g_GangUnit = cdrMillimeter
    g_GangNextJobID = 1

    ' Media defaults: A4 (210 x 297 mm) -- a reasonable starting point
    ' for a gang sheet. These can be overridden by a form later.
    g_GangMediaWidth = 210
    g_GangMediaHeight = 297

    ' Margins: 5mm each side (same as Tile Fill defaults)
    g_GangMarginX = 5
    g_GangMarginY = 5

    ' Gutters: 2mm between copies (same as Tile Fill defaults)
    g_GangGutterX = 2
    g_GangGutterY = 2

    On Error Resume Next
    g_GangUnit = CInt(GetSetting(APP_NAME, SECTION_NAME, "Unit", CStr(g_GangUnit)))
    g_GangMediaWidth = CDbl(GetSetting(APP_NAME, SECTION_NAME, "MediaWidth", CStr(g_GangMediaWidth)))
    g_GangMediaHeight = CDbl(GetSetting(APP_NAME, SECTION_NAME, "MediaHeight", CStr(g_GangMediaHeight)))
    g_GangMarginX = CDbl(GetSetting(APP_NAME, SECTION_NAME, "MarginX", CStr(g_GangMarginX)))
    g_GangMarginY = CDbl(GetSetting(APP_NAME, SECTION_NAME, "MarginY", CStr(g_GangMarginY)))
    g_GangGutterX = CDbl(GetSetting(APP_NAME, SECTION_NAME, "GutterX", CStr(g_GangGutterX)))
    g_GangGutterY = CDbl(GetSetting(APP_NAME, SECTION_NAME, "GutterY", CStr(g_GangGutterY)))
    Err.Clear
    On Error GoTo 0

    g_GangSettingsInitialized = True

End Sub

Sub SaveCurrentGangSettings()
    SaveSetting APP_NAME, SECTION_NAME, "Unit", CStr(g_GangUnit)
    SaveSetting APP_NAME, SECTION_NAME, "MediaWidth", CStr(g_GangMediaWidth)
    SaveSetting APP_NAME, SECTION_NAME, "MediaHeight", CStr(g_GangMediaHeight)
    SaveSetting APP_NAME, SECTION_NAME, "MarginX", CStr(g_GangMarginX)
    SaveSetting APP_NAME, SECTION_NAME, "MarginY", CStr(g_GangMarginY)
    SaveSetting APP_NAME, SECTION_NAME, "GutterX", CStr(g_GangGutterX)
    SaveSetting APP_NAME, SECTION_NAME, "GutterY", CStr(g_GangGutterY)
End Sub
