Attribute VB_Name = "modSettings"
'==============================================================
' modSettings
' Holds all imposition settings as Public variables so both the
' UserForm (frmImpositionSettings) and the main macro
' (modImposition) can read/write them. Defaults live in
' InitDefaultSettings — edit those if you want different
' out-of-the-box values without touching the form.
'
' Last-used persistence: LoadSavedSettings / SaveCurrentSettings use
' VBA's SaveSetting/GetSetting (HKCU, no admin rights, no files).
' The form calls SaveCurrentSettings only after ValidateInputs passes;
' InitDefaultSettings overlays saved values on top of the defaults.
'==============================================================

Option Explicit
Option Private Module

' Registry location for last-used settings. Keep stable: changing these
' orphans previously saved values (which then fall back to defaults).
Private Const SETTINGS_APP As String = "MacrixImposition"
Private Const SETTINGS_SECTION As String = "Settings"

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

    ' Overlay last successfully applied values, if any. Missing or
    ' corrupt entries keep the defaults set above.
    LoadSavedSettings

    g_SettingsInitialized = True
End Sub

'--------------------------------------------------------------
' Persists the current globals as "last successfully applied".
' Saves every setting; a single failing key is reported to the
' Immediate window and skipped so it cannot skip the rest.
' Never raises to the caller (the imposition must always run).
'--------------------------------------------------------------
Public Sub SaveCurrentSettings()
    On Error GoTo SaveError
    SaveSetting SETTINGS_APP, SETTINGS_SECTION, "LayoutMode", g_LayoutMode
    SaveSetting SETTINGS_APP, SETTINGS_SECTION, "Unit", UnitToPersistenceString(g_Unit)
    SaveSetting SETTINGS_APP, SETTINGS_SECTION, "SheetWidth", DoubleToPersistenceString(g_SheetWidth)
    SaveSetting SETTINGS_APP, SETTINGS_SECTION, "SheetHeight", DoubleToPersistenceString(g_SheetHeight)
    SaveSetting SETTINGS_APP, SETTINGS_SECTION, "GridRows", CStr(g_GridRows)
    SaveSetting SETTINGS_APP, SETTINGS_SECTION, "GridCols", CStr(g_GridCols)
    SaveSetting SETTINGS_APP, SETTINGS_SECTION, "GutterX", DoubleToPersistenceString(g_GutterX)
    SaveSetting SETTINGS_APP, SETTINGS_SECTION, "GutterY", DoubleToPersistenceString(g_GutterY)
    SaveSetting SETTINGS_APP, SETTINGS_SECTION, "MarginLeft", DoubleToPersistenceString(g_MarginLeft)
    SaveSetting SETTINGS_APP, SETTINGS_SECTION, "MarginTop", DoubleToPersistenceString(g_MarginTop)
    SaveSetting SETTINGS_APP, SETTINGS_SECTION, "BleedSize", DoubleToPersistenceString(g_BleedSize)
    If g_AddCropMarks Then
        SaveSetting SETTINGS_APP, SETTINGS_SECTION, "AddCropMarks", "1"
    Else
        SaveSetting SETTINGS_APP, SETTINGS_SECTION, "AddCropMarks", "0"
    End If
    SaveSetting SETTINGS_APP, SETTINGS_SECTION, "CropMarkLength", DoubleToPersistenceString(g_CropMarkLength)
    SaveSetting SETTINGS_APP, SETTINGS_SECTION, "StepRepeatPage", CStr(g_StepRepeatPage)
    SaveSetting SETTINGS_APP, SETTINGS_SECTION, "StepRepeatSheetCount", CStr(g_StepRepeatSheetCount)
    Exit Sub
SaveError:
    Debug.Print "MacrixImposition SaveCurrentSettings failed. Error " & Err.Number & " - " & Err.Description
    Err.Clear
    Resume Next
End Sub

'--------------------------------------------------------------
' Overlays persisted values onto the globals. Every entry is read,
' existence-checked and range-validated; anything missing or invalid
' keeps the caller's current value (the built-in defaults when called
' from InitDefaultSettings). Never raises, never shows UI.
'--------------------------------------------------------------
Public Sub LoadSavedSettings()
    On Error Resume Next
    Dim s As String

    s = GetSetting(SETTINGS_APP, SETTINGS_SECTION, "LayoutMode", "")
    If s = "Simple" Or s = "StepRepeat" Or s = "Signature" Then
        g_LayoutMode = s
    ElseIf s <> "" Then
        Debug.Print "MacrixImposition: ignoring invalid saved 'LayoutMode'=[" & s & "], keeping " & g_LayoutMode
    End If

    s = GetSetting(SETTINGS_APP, SETTINGS_SECTION, "Unit", "")
    If s = "Millimeters" Then
        g_Unit = cdrMillimeter
    ElseIf s = "Inches" Then
        g_Unit = cdrInch
    ElseIf s = "Points" Then
        g_Unit = cdrPoint
    ElseIf s = "Pixels" Then
        g_Unit = cdrPixel
    ElseIf s <> "" Then
        Debug.Print "MacrixImposition: ignoring invalid saved 'Unit'=[" & s & "]"
    End If

    g_SheetWidth = GetSavedPositiveDouble("SheetWidth", g_SheetWidth)
    g_SheetHeight = GetSavedPositiveDouble("SheetHeight", g_SheetHeight)
    g_GridRows = GetSavedPositiveWholeInt("GridRows", g_GridRows)
    g_GridCols = GetSavedPositiveWholeInt("GridCols", g_GridCols)
    g_GutterX = GetSavedNonNegativeDouble("GutterX", g_GutterX)
    g_GutterY = GetSavedNonNegativeDouble("GutterY", g_GutterY)
    g_MarginLeft = GetSavedNonNegativeDouble("MarginLeft", g_MarginLeft)
    g_MarginTop = GetSavedNonNegativeDouble("MarginTop", g_MarginTop)
    g_BleedSize = GetSavedNonNegativeDouble("BleedSize", g_BleedSize)
    g_CropMarkLength = GetSavedNonNegativeDouble("CropMarkLength", g_CropMarkLength)
    g_StepRepeatPage = GetSavedPositiveWholeInt("StepRepeatPage", g_StepRepeatPage)
    g_StepRepeatSheetCount = GetSavedPositiveWholeInt("StepRepeatSheetCount", g_StepRepeatSheetCount)

    s = GetSetting(SETTINGS_APP, SETTINGS_SECTION, "AddCropMarks", "")
    If s = "1" Or StrComp(s, "True", vbTextCompare) = 0 Then
        g_AddCropMarks = True
    ElseIf s = "0" Or StrComp(s, "False", vbTextCompare) = 0 Then
        g_AddCropMarks = False
    ElseIf s <> "" Then
        Debug.Print "MacrixImposition: ignoring invalid saved 'AddCropMarks'=[" & s & "]"
    End If

    Err.Clear
    On Error GoTo 0

    Debug.Print "MacrixImposition settings loaded: LayoutMode=" & g_LayoutMode & _
        " Unit=" & UnitToPersistenceString(g_Unit) & _
        " Sheet=" & DoubleToPersistenceString(g_SheetWidth) & "x" & DoubleToPersistenceString(g_SheetHeight) & _
        " Grid=" & g_GridRows & "x" & g_GridCols & _
        " Gutter=" & DoubleToPersistenceString(g_GutterX) & "/" & DoubleToPersistenceString(g_GutterY) & _
        " Margin=" & DoubleToPersistenceString(g_MarginLeft) & "/" & DoubleToPersistenceString(g_MarginTop) & _
        " Bleed=" & DoubleToPersistenceString(g_BleedSize) & _
        " Crop=" & g_AddCropMarks & "/" & DoubleToPersistenceString(g_CropMarkLength) & _
        " StepRepeat=" & g_StepRepeatPage & "/" & g_StepRepeatSheetCount
End Sub

'--------------------------------------------------------------
' Helpers: read one registry value, validate it, else Fallback.
' A present-but-invalid value is reported to the Immediate window;
' a missing value (first run) stays silent.
'--------------------------------------------------------------
Private Function GetSavedPositiveDouble(ByVal Key As String, ByVal Fallback As Double) As Double
    On Error GoTo LoadError
    Dim raw As String
    Dim v As Double
    raw = GetSetting(SETTINGS_APP, SETTINGS_SECTION, Key, "")
    If raw = "" Then
        GetSavedPositiveDouble = Fallback
        Exit Function
    End If
    If TryParsePersistedDouble(raw, v) And v > 0 Then
        GetSavedPositiveDouble = v
    Else
        Debug.Print "MacrixImposition: ignoring invalid saved '" & Key & "'=[" & raw & "], keeping " & Fallback
        GetSavedPositiveDouble = Fallback
    End If
    Exit Function
LoadError:
    Debug.Print "MacrixImposition: error reading saved '" & Key & "'. Error " & Err.Number & " - " & Err.Description
    Err.Clear
    GetSavedPositiveDouble = Fallback
End Function

Private Function GetSavedNonNegativeDouble(ByVal Key As String, ByVal Fallback As Double) As Double
    On Error GoTo LoadError
    Dim raw As String
    Dim v As Double
    raw = GetSetting(SETTINGS_APP, SETTINGS_SECTION, Key, "")
    If raw = "" Then
        GetSavedNonNegativeDouble = Fallback
        Exit Function
    End If
    If TryParsePersistedDouble(raw, v) And v >= 0 Then
        GetSavedNonNegativeDouble = v
    Else
        Debug.Print "MacrixImposition: ignoring invalid saved '" & Key & "'=[" & raw & "], keeping " & Fallback
        GetSavedNonNegativeDouble = Fallback
    End If
    Exit Function
LoadError:
    Debug.Print "MacrixImposition: error reading saved '" & Key & "'. Error " & Err.Number & " - " & Err.Description
    Err.Clear
    GetSavedNonNegativeDouble = Fallback
End Function

Private Function GetSavedPositiveWholeInt(ByVal Key As String, ByVal Fallback As Integer) As Integer
    On Error GoTo LoadError
    Dim raw As String
    Dim v As Double
    raw = GetSetting(SETTINGS_APP, SETTINGS_SECTION, Key, "")
    If raw = "" Then
        GetSavedPositiveWholeInt = Fallback
        Exit Function
    End If
    If TryParsePersistedDouble(raw, v) And v >= 1 And v <= 32767 And v = Int(v) Then
        GetSavedPositiveWholeInt = CInt(v)
    Else
        Debug.Print "MacrixImposition: ignoring invalid saved '" & Key & "'=[" & raw & "], keeping " & Fallback
        GetSavedPositiveWholeInt = Fallback
    End If
    Exit Function
LoadError:
    Debug.Print "MacrixImposition: error reading saved '" & Key & "'. Error " & Err.Number & " - " & Err.Description
    Err.Clear
    GetSavedPositiveWholeInt = Fallback
End Function

' Parses a persisted decimal. Accepts our dot-invariant format
' ("450", "17.7165", ".5", "1E+20") as well as legacy values written
' with the Windows locale decimal separator. Returns True + value,
' or False (outVal untouched).
Private Function TryParsePersistedDouble(ByVal raw As String, ByRef outVal As Double) As Boolean
    On Error GoTo ParseError
    Dim t As String
    Dim sep As String
    t = Trim(raw)
    If t = "" Then Exit Function
    sep = Mid(CStr(1.5), 2, 1)
    If sep <> "." Then t = Replace(t, ".", sep)
    If IsNumeric(t) Then
        outVal = CDbl(t)
        If Err.Number = 0 Then
            TryParsePersistedDouble = True
            Exit Function
        End If
    End If
    Exit Function
ParseError:
    Debug.Print "MacrixImposition: cannot parse persisted number [" & raw & "]. Error " & Err.Number & " - " & Err.Description
    Err.Clear
End Function

' Serializes a Double locale-independently (always a dot separator).
Private Function DoubleToPersistenceString(ByVal v As Double) As String
    DoubleToPersistenceString = Trim(Str(v))
End Function

' Stable unit serialization: names, never raw enum numerics.
Private Function UnitToPersistenceString(ByVal u As cdrUnit) As String
    Select Case u
        Case cdrMillimeter: UnitToPersistenceString = "Millimeters"
        Case cdrInch:       UnitToPersistenceString = "Inches"
        Case cdrPoint:      UnitToPersistenceString = "Points"
        Case cdrPixel:      UnitToPersistenceString = "Pixels"
        Case Else:          UnitToPersistenceString = "Millimeters"
    End Select
End Function

'--------------------------------------------------------------
' Diagnostic: proves SaveSetting/GetSetting round-trips in this host.
' Run from CorelDRAW (Macros > Run Macro). Must show HELLO123, and
' still show it after a full CorelDRAW restart. Any other result
' means registry persistence itself is unavailable: diagnose that
' before touching the form or the init flow.
'--------------------------------------------------------------
Public Sub TestSettingsPersistence()

    Dim testValue As String

    On Error GoTo TestError

    SaveSetting "MacrixImposition", "Settings", "PersistenceTest", "HELLO123"

    testValue = GetSetting( _
        "MacrixImposition", _
        "Settings", _
        "PersistenceTest", _
        "<MISSING>" _
    )

    MsgBox "Saved/read value = [" & testValue & "]", _
           vbInformation, _
           "Persistence Test"

    Exit Sub

TestError:

    MsgBox _
        "Persistence test failed." & vbCrLf & _
        "Error: " & Err.Number & vbCrLf & _
        "Description: " & Err.Description, _
        vbCritical, _
        "Persistence Test"

End Sub

'--------------------------------------------------------------
' Diagnostic: dumps every persisted imposition value currently in
' the registry (<missing> = never saved). Use it to decide which
' side is broken: keys absent -> SaveCurrentSettings never ran or
' failed; keys present but form shows defaults -> LoadSavedSettings
' / init flow / stale generated form code.
'--------------------------------------------------------------
Public Sub DebugShowSavedSettings()

    Dim s As String

    On Error GoTo ShowError

    s = "LayoutMode=" & GetSetting(SETTINGS_APP, SETTINGS_SECTION, "LayoutMode", "<missing>") & vbCrLf
    s = s & "Unit=" & GetSetting(SETTINGS_APP, SETTINGS_SECTION, "Unit", "<missing>") & vbCrLf
    s = s & "SheetWidth=" & GetSetting(SETTINGS_APP, SETTINGS_SECTION, "SheetWidth", "<missing>") & vbCrLf
    s = s & "SheetHeight=" & GetSetting(SETTINGS_APP, SETTINGS_SECTION, "SheetHeight", "<missing>") & vbCrLf
    s = s & "GridRows=" & GetSetting(SETTINGS_APP, SETTINGS_SECTION, "GridRows", "<missing>") & vbCrLf
    s = s & "GridCols=" & GetSetting(SETTINGS_APP, SETTINGS_SECTION, "GridCols", "<missing>") & vbCrLf
    s = s & "GutterX=" & GetSetting(SETTINGS_APP, SETTINGS_SECTION, "GutterX", "<missing>") & vbCrLf
    s = s & "GutterY=" & GetSetting(SETTINGS_APP, SETTINGS_SECTION, "GutterY", "<missing>") & vbCrLf
    s = s & "MarginLeft=" & GetSetting(SETTINGS_APP, SETTINGS_SECTION, "MarginLeft", "<missing>") & vbCrLf
    s = s & "MarginTop=" & GetSetting(SETTINGS_APP, SETTINGS_SECTION, "MarginTop", "<missing>") & vbCrLf
    s = s & "BleedSize=" & GetSetting(SETTINGS_APP, SETTINGS_SECTION, "BleedSize", "<missing>") & vbCrLf
    s = s & "AddCropMarks=" & GetSetting(SETTINGS_APP, SETTINGS_SECTION, "AddCropMarks", "<missing>") & vbCrLf
    s = s & "CropMarkLength=" & GetSetting(SETTINGS_APP, SETTINGS_SECTION, "CropMarkLength", "<missing>") & vbCrLf
    s = s & "StepRepeatPage=" & GetSetting(SETTINGS_APP, SETTINGS_SECTION, "StepRepeatPage", "<missing>") & vbCrLf
    s = s & "StepRepeatSheetCount=" & GetSetting(SETTINGS_APP, SETTINGS_SECTION, "StepRepeatSheetCount", "<missing>")

    MsgBox s, vbInformation, "Saved Macrix Settings"

    Exit Sub

ShowError:

    MsgBox _
        "Could not read saved settings." & vbCrLf & _
        "Error: " & Err.Number & vbCrLf & _
        "Description: " & Err.Description, _
        vbCritical, _
        "Saved Macrix Settings"

End Sub
