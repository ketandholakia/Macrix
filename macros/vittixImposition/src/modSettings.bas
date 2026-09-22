'==============================================================
' modSettings
' Holds all imposition settings as Public variables so both the
' UserForm (frmImpositionSettings) and the main macro
' (modImposition) can read/write them. Defaults live in
' InitDefaultSettings — edit those if you want different
' out-of-the-box values without touching the form.
'==============================================================

Option Explicit

Public g_SettingsInitialized As Boolean

Public g_Unit As cdrUnit

' "Simple", "StepRepeat", or "Signature"
Public g_LayoutMode As String

Public g_SheetWidth As Double
Public g_SheetHeight As Double

Public g_GridRows As Integer
Public g_GridCols As Integer

Public g_GutterX As Double
Public g_GutterY As Double

Public g_MarginLeft As Double
Public g_MarginTop As Double

Public g_BleedSize As Double
Public g_AddCropMarks As Boolean
Public g_CropMarkLength As Double

' Only used when g_LayoutMode = "StepRepeat"
Public g_StepRepeatPage As Integer
Public g_StepRepeatSheetCount As Integer

Sub InitDefaultSettings()
    g_Unit = cdrMillimeter
    g_LayoutMode = "Simple"

    g_SheetWidth = 320
    g_SheetHeight = 450

    g_GridRows = 2
    g_GridCols = 2

    g_GutterX = 5
    g_GutterY = 5

    g_MarginLeft = 10
    g_MarginTop = 10

    g_BleedSize = 3
    g_AddCropMarks = True
    g_CropMarkLength = 5

    g_StepRepeatPage = 1
    g_StepRepeatSheetCount = 4

    g_SettingsInitialized = True
End Sub
