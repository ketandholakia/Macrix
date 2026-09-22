Attribute VB_Name = "modAppServices"
Option Explicit
''
' modAppServices — Application-level utilities (optimization, events, command
' groups, active doc/selection) on the CorelDRAW host.
'
' Binding: late-bound by policy (§42). Callers pass the host `Application`
' object so nothing here hard-locks to a type-library version.
'
' Undo/reentrancy (§14, §39): callers MUST pair SV_App_SetOptimization(False) /
' SetEventsEnabled(False) with a restore on EVERY path (single cleanup label).
' This module only exposes the toggles; the pairing contract lives with the caller.
''
' Depends on: modComLifecycle, modLogger.
' CorelDRAW: yes — characterization-test manually (§40/41), never auto-verified.

' Resolve the host Application object (global provided by CorelDRAW).
Public Function SV_App_Host() As Object
    On Error Resume Next
    Set SV_App_Host = Application
    On Error GoTo 0
End Function

Public Function SV_App_ActiveDocument(ByVal app As Object) As Object
    On Error Resume Next
    Set SV_App_ActiveDocument = CallByName(app, "ActiveDocument", VbGet)
    On Error GoTo 0
End Function

Public Function SV_App_ActiveSelection(ByVal app As Object) As Object
    On Error Resume Next
    Set SV_App_ActiveSelection = CallByName(app, "ActiveSelection", VbGet)
    On Error GoTo 0
End Function

Public Sub SV_App_SetOptimization(ByVal app As Object, ByVal flag As Boolean)
    On Error Resume Next
    CallByName app, "Optimization", VbLet, flag
    On Error GoTo 0
End Sub

Public Sub SV_App_SetEventsEnabled(ByVal app As Object, ByVal flag As Boolean)
    On Error Resume Next
    CallByName app, "EventsEnabled", VbLet, flag
    On Error GoTo 0
End Sub

' Best-effort: start/stop a command (undo) group; the member name should be
' confirmed against the installed type library (BeginCommandGroup/EndCommandGroup).
Public Sub SV_App_BeginCommandGroup(ByVal app As Object)
    On Error Resume Next
    CallByName app, "BeginCommandGroup", VbMethod
    On Error GoTo 0
End Sub

Public Sub SV_App_EndCommandGroup(ByVal app As Object)
    On Error Resume Next
    CallByName app, "EndCommandGroup", VbMethod
    On Error GoTo 0
End Sub