Attribute VB_Name = "modPageServices"
Option Explicit
''
' modPageServices — Page-level utilities (count, size, active). Late-bound.
' Depends on: modComLifecycle. CorelDRAW: yes (characterization).

Public Function SV_Pg_ActivePage(ByVal doc As Object) As Object
    On Error Resume Next
    Set SV_Pg_ActivePage = CallByName(doc, "ActivePage", VbGet)
    On Error GoTo 0
End Function

Public Function SV_Pg_Width(ByVal page As Object) As Double
    On Error Resume Next
    SV_Pg_Width = CDbl(CallByName(page, "SizeWidth", VbGet))   ' confirm member name
    On Error GoTo 0
End Function