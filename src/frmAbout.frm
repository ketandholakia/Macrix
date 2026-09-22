VERSION 5.0
Begin VB.UserForm frmAbout
   Caption         =   "Vittix CorelDRAW Macros"
   ClientHeight    =   2280
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4680
   StartUpPosition =   1
   Begin VB.CommandButton cmdClose
      Caption         =   "Close"
      Height          =   405
      Left            =   3720
      TabIndex        =   1
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label lblInfo
      Caption         =   "Vittix CorelDRAW Macros"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   180
      Width           =   4155
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_Exposed = False
Option Explicit
''
' frmAbout — About dialog. Depends on modVersion. CorelDRAW: none.

Public Sub OpenDialog()
    lblInfo.Caption = modVersion.VER_KitName & " v" & modVersion.VER_CurrentVersion & vbCrLf & _
        "Copyright (c) 2026 Vittix"
    modDialogHelper.DH_ShowModal Me
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub