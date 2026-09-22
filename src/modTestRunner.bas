Attribute VB_Name = "modTestRunner"
Option Explicit
''
' modTestRunner — lightweight unit-test harness (§40).
'
' Runs the pure-logic modules (pl_*) and prints pass/fail to the Immediate window
' plus the diagnostics log. There is no third-party framework; this runner asserts
' via T_Check and keeps a running tally. CorelDRAW-dependent behavior is covered
' separately by the manual characterization suite (tests/manual).
'
' Usage (CorelDRAW VBA Immediate or a macro):  modTestRunner.RunAll
'
' Depends on: pl_StringUtils, pl_GeometryUtils, modVersion, modVersionDetect,
' modFeatureRegistry, modLogger.
' CorelDRAW: none (pure). modFeatureRegistry/modVersionDetect are exercised
' here only through their pure, non-COM entry points.

Public Sub RunAll()
    Dim pass As Long, fail As Long, started As Double
    started = Timer
    modLogger.LOG_Info "TEST RUN begin"

    T_Check pl_StringUtils.PL_WordCount("a b c") = 3, pass, fail, "WordCount=3"
    T_Check pl_StringUtils.PL_WordCount("") = 0, pass, fail, "WordCount empty=0"
    T_Check pl_StringUtils.PL_StripDigits("a1b2c3") = "abc", pass, fail, "StripDigits a1b2c3"
    T_Check pl_StringUtils.PL_PadLeft("x", 3, "0") = "00x", pass, fail, "PadLeft 00x"

    T_Check pl_GeometryUtils.PLG_Clamp(5, 0, 3) = 3, pass, fail, "Clamp hi"
    T_Check pl_GeometryUtils.PLG_Clamp(-1, 0, 3) = 0, pass, fail, "Clamp lo"
    T_Check pl_GeometryUtils.PLG_Clamp(2, 0, 3) = 2, pass, fail, "Clamp mid"
    T_Check pl_GeometryUtils.PLG_Distance(0, 0, 3, 4) = 5, pass, fail, "3-4-5 triangle"

    ' modVersionDetect.VD_ParseCorelVersion: segments must sort independently
    ' of digit-count (this regressed once already — see CHANGELOG).
    T_Check modVersionDetect.VD_ParseCorelVersion("2021.0.0") > _
        modVersionDetect.VD_ParseCorelVersion("18.0.0.448"), pass, fail, "ParseCorelVersion 2021>18.x build"
    T_Check modVersionDetect.VD_ParseCorelVersion("9.0.0") < _
        modVersionDetect.VD_ParseCorelVersion("10.0.0"), pass, fail, "ParseCorelVersion 9<10"
    T_Check modVersionDetect.VD_ParseCorelVersion("") = 0, pass, fail, "ParseCorelVersion empty=0"

    ' modFeatureRegistry.FR_RegisterFeatures must be idempotent (safe to call
    ' more than once per session without duplicating entries).
    modFeatureRegistry.FR_RegisterFeatures
    Dim countAfterFirst As Long
    countAfterFirst = modFeatureRegistry.FR_Count()
    modFeatureRegistry.FR_RegisterFeatures
    T_Check modFeatureRegistry.FR_Count() = countAfterFirst, pass, fail, "FR_RegisterFeatures idempotent"

    modLogger.LOG_Info "TEST RUN end: pass=" & CStr(pass) & " fail=" & CStr(fail)
    Debug.Print "Vittix tests: pass=" & pass & " fail=" & fail
End Sub

' Record a single assertion; returns True when the assertion holds.
Public Function T_Check(ByVal ok As Boolean, ByRef pass As Long, ByRef fail As Long, _
                        ByVal label As String) As Boolean
    If ok Then
        pass = pass + 1
        T_Check = True
        Debug.Print "  PASS  " & label
    Else
        fail = fail + 1
        modLogger.LOG_Warn "  FAIL  " & label
        Debug.Print "  FAIL  " & label
    End If
End Function