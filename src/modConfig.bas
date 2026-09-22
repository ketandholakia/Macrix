Attribute VB_Name = "modConfig"
Option Explicit
''
' modConfig: central configuration/settings (§8).
' All knobs resolve through Property Get so defaults and overrides are
' traceable. No external file dependency by default; an optional persisted
' settings path is handled via modPathManager if a settings store is added.
''
' Depends on: modPathManager (for the data dir).
' Pure/logic: the getters are deterministic; no Corel object model.

Public Enum CFG_Levels
    CFG_LevelDebug = 0
    CFG_LevelInfo = 1
    CFG_LevelWarn = 2
    CFG_LevelError = 3
End Enum

Private m_LogLevel As Long
Private m_LogToFile As Boolean
Private m_IsInitialized As Boolean

Public Property Get CFG_LogLevel() As Long
    If Not m_IsInitialized Then CFG_Initialize
    CFG_LogLevel = m_LogLevel
End Property

Public Property Get CFG_LogToFile() As Boolean
    If Not m_IsInitialized Then CFG_Initialize
    CFG_LogToFile = m_LogToFile
End Property

' Circuit breaker tolerance: failures within one session before a feature
' auto-disables for the remainder of the session (§46).
Public Property Get CFG_CircuitBreakerThreshold() As Long
    CFG_CircuitBreakerThreshold = 3
End Property

' Opt-in telemetry default is OFF; nothing is transmitted anywhere (§44).
Public Property Get CFG_TelemetryAllowed() As Boolean
    CFG_TelemetryAllowed = False
End Property

Private Sub CFG_Initialize()
    m_LogLevel = CFG_LevelInfo
    m_LogToFile = True
    m_IsInitialized = True
End Sub

' Public diagnostics dump for the About/log header.
Public Function CFG_Describe() As String
    CFG_Describe = "logLevel=" & CStr(m_LogLevel) & ", logToFile=" & CStr(m_LogToFile) & ", telemetry=off"
End Function