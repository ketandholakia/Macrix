Attribute VB_Name = "pl_StringUtils"
Option Explicit
''
' pl_StringUtils — pure string helpers with NO CorelDRAW object-model references.
' Unit-tested by modTestRunner (§40). These must never touch COM.
'
' Bad words / digits: keep this module pure so tests run anywhere.

' Return the number of whitespace-separated words.
Public Function PL_WordCount(ByVal text As String) As Long
    Dim t As String, cnt As Long, inWord As Boolean, i As Long
    t = text
    cnt = 0: inWord = False
    For i = 1 To Len(t)
        If Mid$(t, i, 1) = " " Or Mid$(t, i, 1) = vbTab Or Mid$(t, i, 1) = vbCr Or Mid$(t, i, 1) = vbLf Then
            inWord = False
        Else
            If Not inWord Then cnt = cnt + 1
            inWord = True
        End If
    Next i
    PL_WordCount = cnt
End Function

' Remove every digit (0-9) from `text`. Pure.
Public Function PL_StripDigits(ByVal text As String) As String
    Dim res As String, i As Long, ch As String
    res = ""
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If ch < "0" Or ch > "9" Then res = res & ch
    Next i
    PL_StripDigits = res
End Function

' Left-pad to `length` with `pad` (repeated). Returns input unchanged if longer.
Public Function PL_PadLeft(ByVal text As String, ByVal length As Long, Optional ByVal pad As String = " ") As String
    If pad = "" Then pad = " "
    Dim result As String
    result = text
    While Len(result) < length
        result = pad & result
    Wend
    PL_PadLeft = result
End Function