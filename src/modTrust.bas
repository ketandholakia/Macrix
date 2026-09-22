Attribute VB_Name = "modTrust"
Option Explicit
''
' modTrust — macro security, signing & trust guidance (§38).
'
' Policy defaults (documented assumptions — owner inputs pending):
'   • The macro project is assumed UNSIGNED for now (no cert supplied).
'   • Therefore the installer must tell the user, in plain language, that they
'     must trust the macro project on first run (CorelDRAW VBA security).
'   • We do NOT password-protect the project as a substitute for real access
'     control; if a password is ever added it is to prevent accidental edits.
''
' Depends on: nothing.
' CorelDRAW dependency: none (policy + messaging).

Public Function TRUST_HasSignature() As Boolean
    ' Bump to True when a cert is actually attached + documented (name, expiry,
    ' renewal) per §38 — otherwise the renewal-day trust break is silent.
    TRUST_HasSignature = False
End Function

Public Function TRUST_RequiresManualTrust() As Boolean
    TRUST_RequiresManualTrust = Not TRUST_HasSignature()
End Function

' Plain-language copy the installer/About box can surface.
Public Function TRUST_FirstRunMessage() As String
    If TRUST_HasSignature() Then
        TRUST_FirstRunMessage = ""
    Else
        TRUST_FirstRunMessage = _
            "This macro project is unsigned. On first run, CorelDRAW may ask you to " & _
            "enable macros for this project in VBA security. Only do so if you trust " & _
            "the source of this file."
    End If
End Function

' Whether the project is locked with a password (documented reason only).
Public Function TRUST_ProjectPasswordProtected() As Boolean
    TRUST_ProjectPasswordProtected = False
End Function

' Certificate metadata for the docs when signed.
Public Function TRUST_CertSubject() As String
    TRUST_CertSubject = ""
End Function