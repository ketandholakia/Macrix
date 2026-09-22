Attribute VB_Name = "modBootstrap"
Option Explicit
''
' modBootstrap — entry / bootstrap orchestration (§4 Bootstrap/Entry Points).
'
' Main() is the canonical start: it (re)initializes the reference registry and
' feature registry and logs a start banner with version + diagnostics header. It
' does NOT auto-run features; user-facing entry points are the individual feature
' modules invoked from CorelDRAW, each wrapped in the safe-run pattern.
''
' Depends on: modReferenceRegistry, modFeatureRegistry, modLogger, modVersion,
' modConfig.
' CorelDRAW dependency: none.

Public Const BST_ModuleName As String = "modBootstrap"

Public Sub Main()
    modReferenceRegistry.REF_Initialize
    modFeatureRegistry.FR_RegisterFeatures
    modLogger.LOG_Info modVersion.VER_KitName & " v" & modVersion.VER_CurrentVersion & _
        " started (" & modConfig.CFG_Describe() & ")"
    modLogger.LOG_Info "features: " & modFeatureRegistry.FR_Describe()
    modLogger.LOG_Info "references: " & modReferenceRegistry.REF_AllNames()
End Sub

' Convenience "i'm here" action for testing the bootstrap chain without a doc.
Public Sub Ping()
    Debug.Print modVersion.VER_KitName & " alive at " & Now
End Sub