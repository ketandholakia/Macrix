VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsGangJob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'==============================================================
' clsGangJob
' One production job in the Gang Job list: a rectangular piece of
' artwork (Width x Height) needed in some Quantity. Phase 1 data
' model only -- no CorelDRAW output yet.
'
' Width/Height are read once from the source selection's bounding
' box (ShapeRange.SizeWidth/SizeHeight) when the job is added, in
' modGangJobSettings.g_GangUnit units, and are NOT re-read later --
' if the source artwork changes size afterwards, re-add the job.
'
' This is the project's first class module (everything else here
' is plain standard modules) -- it exists because a VBA Collection
' cannot hold user-defined Types directly; a job needs distinct,
' independently-addressable instances, so a class is the smallest
' change that gets there.
'==============================================================

Public JobID As Long
Public JobName As String

Public Width As Double
Public Height As Double
Public Quantity As Long

Public AllowRotation As Boolean
