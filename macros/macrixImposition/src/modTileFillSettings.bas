Attribute VB_Name = "modTileFillSettings"
'==============================================================
' modTileFillSettings
' Settings for modTileFill (RunTileFill). Public "g_Tile..."
' variables so both RunTileFill and frmTileFillSettings can read
' and write them.
'
' Persistence: SaveCurrentTileSettings/last-used overlay uses the
' standard VBA SaveSetting/GetSetting registry helpers (HKCU\Software
' \VB and VBA Program Settings\macrixTileFill\Settings). If
' modSettings.bas's SaveCurrentSettings uses a different scheme
' (shared registry root, a helper module, etc.), point this at that
' instead so both macros persist the same way.
'==============================================================

Option Explicit
Option Private Module

Private Const APP_NAME As String = "macrixTileFill"
Private Const SECTION_NAME As String = "Settings"

Public g_TileSettingsInitialized As Boolean

Public g_TileUnit As cdrUnit

' Minimum gap kept between the tiled grid and the page edge.
Public g_TileMarginX As Double
Public g_TileMarginY As Double

' Gap kept between adjacent copies.
Public g_TileGutterX As Double
Public g_TileGutterY As Double

' If True, RunTileFill will also test the selection rotated 90 degrees
' and use whichever orientation fits more copies on the page.
Public g_TileAllowRotate As Boolean

'--------------------------------------------------------------
' Resets to built-in defaults, then overlays last-used values from
' the registry if any were saved by a previous successful run.
' Guarded (like InitDefaultSettings) so calling this again mid-
' session never wipes in-memory values back to the hard defaults.
'--------------------------------------------------------------
Sub InitDefaultTileSettings()

    g_TileUnit = cdrMillimeter
    g_TileMarginX = 5
    g_TileMarginY = 5
    g_TileGutterX = 2
    g_TileGutterY = 2
    g_TileAllowRotate = True

    On Error Resume Next
    g_TileUnit = CInt(GetSetting(APP_NAME, SECTION_NAME, "Unit", CStr(g_TileUnit)))
    g_TileMarginX = CDbl(GetSetting(APP_NAME, SECTION_NAME, "MarginX", CStr(g_TileMarginX)))
    g_TileMarginY = CDbl(GetSetting(APP_NAME, SECTION_NAME, "MarginY", CStr(g_TileMarginY)))
    g_TileGutterX = CDbl(GetSetting(APP_NAME, SECTION_NAME, "GutterX", CStr(g_TileGutterX)))
    g_TileGutterY = CDbl(GetSetting(APP_NAME, SECTION_NAME, "GutterY", CStr(g_TileGutterY)))
    g_TileAllowRotate = CBool(GetSetting(APP_NAME, SECTION_NAME, "AllowRotate", CStr(g_TileAllowRotate)))
    Err.Clear
    On Error GoTo 0

    g_TileSettingsInitialized = True

End Sub

'--------------------------------------------------------------
' Persists the current settings so the next session's
' InitDefaultTileSettings picks them back up. Called by
' frmTileFillSettings only after a successful RunTileFill, same
' as frmImpositionSettings does with SaveCurrentSettings.
'--------------------------------------------------------------
Sub SaveCurrentTileSettings()
    SaveSetting APP_NAME, SECTION_NAME, "Unit", CStr(g_TileUnit)
    SaveSetting APP_NAME, SECTION_NAME, "MarginX", CStr(g_TileMarginX)
    SaveSetting APP_NAME, SECTION_NAME, "MarginY", CStr(g_TileMarginY)
    SaveSetting APP_NAME, SECTION_NAME, "GutterX", CStr(g_TileGutterX)
    SaveSetting APP_NAME, SECTION_NAME, "GutterY", CStr(g_TileGutterY)
    SaveSetting APP_NAME, SECTION_NAME, "AllowRotate", CStr(g_TileAllowRotate)
End Sub

'--------------------------------------------------------------
' Quick InputBox-based prompt, kept as a no-form fallback (e.g. for
' testing from the Immediate Window). frmTileFillSettings is the
' normal way to change these now.
'--------------------------------------------------------------
Sub PromptTileSettings()

    If Not g_TileSettingsInitialized Then InitDefaultTileSettings

    Dim resp As String
    Dim unitLabel As String
    unitLabel = TileUnitName(g_TileUnit)

    resp = InputBox("Margin from page edge, X (" & unitLabel & "):", _
                     "Tile Fill Settings", CStr(g_TileMarginX))
    If resp = "" Then Exit Sub
    g_TileMarginX = Val(resp)

    resp = InputBox("Margin from page edge, Y (" & unitLabel & "):", _
                     "Tile Fill Settings", CStr(g_TileMarginY))
    If resp = "" Then Exit Sub
    g_TileMarginY = Val(resp)

    resp = InputBox("Gutter between copies, X (" & unitLabel & "):", _
                     "Tile Fill Settings", CStr(g_TileGutterX))
    If resp = "" Then Exit Sub
    g_TileGutterX = Val(resp)

    resp = InputBox("Gutter between copies, Y (" & unitLabel & "):", _
                     "Tile Fill Settings", CStr(g_TileGutterY))
    If resp = "" Then Exit Sub
    g_TileGutterY = Val(resp)

    resp = InputBox("Allow 90-degree rotation if it fits more copies? (Y/N):", _
                     "Tile Fill Settings", IIf(g_TileAllowRotate, "Y", "N"))
    If resp = "" Then Exit Sub
    g_TileAllowRotate = (UCase(Left(resp, 1)) = "Y")

    g_TileSettingsInitialized = True

    On Error Resume Next
    SaveCurrentTileSettings
    Err.Clear
    On Error GoTo 0

    MsgBox "Tile Fill settings updated.", vbInformation

End Sub

Private Function TileUnitName(u As cdrUnit) As String
    Select Case u
        Case cdrMillimeter: TileUnitName = "mm"
        Case cdrInch: TileUnitName = "inch"
        Case cdrPoint: TileUnitName = "pt"
        Case cdrPixel: TileUnitName = "px"
        Case Else: TileUnitName = "units"
    End Select
End Function
