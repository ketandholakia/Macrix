Attribute VB_Name = "pl_GeometryUtils"
Option Explicit
''
' pl_GeometryUtils — pure 2D geometry helpers, NO CorelDRAW references.
' Unit-tested by modTestRunner (§40).

' Clamp v into [lo, hi].
Public Function PLG_Clamp(ByVal v As Double, ByVal lo As Double, ByVal hi As Double) As Double
    If v < lo Then PLG_Clamp = lo: Exit Function
    If v > hi Then PLG_Clamp = hi: Exit Function
    PLG_Clamp = v
End Function

Public Function PLG_Distance(ByVal x1 As Double, ByVal y1 As Double, _
                             ByVal x2 As Double, ByVal y2 As Double) As Double
    PLG_Distance = Sqr(((x2 - x1) ^ 2) + ((y2 - y1) ^ 2))
End Function

Public Function PLG_Round(ByVal v As Double, Optional ByVal decimals As Long = 2) As Double
    Dim factor As Double
    factor = 10 ^ decimals
    PLG_Round = Int(v * factor + 0.5) / factor
End Function