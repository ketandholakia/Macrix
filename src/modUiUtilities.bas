Attribute VB_Name = "modUiUtilities"
Option Explicit
''
' modUiUtilities — small UI conveniences (message boxes, confirm, progress).
' Encapsulates user-facing prompts so features never call MsgBox directly.
' Depends on: modVersion. CorelDRAW: none.

Public Sub UU_MsgInfo(ByVal text As String)
    MsgBox text, vbOKOnly Or vbInformation, modVersion.VER_KitName
End Sub

Public Sub UU_MsgWarn(ByVal text As String)
    MsgBox text, vbOKOnly Or vbExclamation, modVersion.VER_KitName
End Sub

Public Sub UU_MsgError(ByVal text As String)
    MsgBox text, vbOKOnly Or vbCritical, modVersion.VER_KitName
End Sub

' Yes/No confirm; default intended response is No.
Public Function UU_Confirm(ByVal question As String) As Boolean
    Dim r As Variant
    r = MsgBox(question, vbYesNo Or vbQuestion Or vbDefaultButton2, modVersion.VER_KitName)
    UU_Confirm = (r = vbYes)
End Function