Attribute VB_Name = "modDialogHelper"
Option Explicit
''
' modDialogHelper — UserForm lifecycle helpers (center/position, load, show).
' Forms are passed late-bound (Object) so this module does not depend on any
' concrete form class. Depends on: nothing. CorelDRAW: none.

Public Sub DH_Center(ByVal frm As Object)
    On Error Resume Next
    Dim x As Long, y As Long
    x = (Screen.Width - frm.Width) \ 2
    y = (Screen.Height - frm.Height) \ 2
    frm.Move x, y, frm.Width, frm.Height
    On Error GoTo 0
End Sub

Public Sub DH_ShowModal(ByVal frm As Object)
    DH_Center frm
    frm.Show vbModal
End Sub