Attribute VB_Name = "Vb_Str_Rmv"
Option Compare Database
Option Explicit

Function RmvDblSpc$(S)
Dim O$: O = S
While HasSubStr(O, "  ")
    O = Replace(O, "  ", " ")
Wend
RmvDblSpc = O
End Function

Function RmvFstChr$(S)
RmvFstChr = RmvFstNChr(S)
End Function

Function RmvFstNChr$(S, Optional N% = 1)
RmvFstNChr = Mid(S, N + 1)
End Function

Function RmvLasChr$(S)
RmvLasChr = RmvLasNChr(S)
End Function

Function RmvLasNChr$(S, Optional N% = 1)
RmvLasNChr = Left(S, Len(S) - 1)
End Function

Function RmvPfx$(S, Pfx)
Dim L%: L = Len(Pfx)
If Left(S, L) = Pfx Then
    RmvPfx = Mid(S, L + 1)
Else
    RmvPfx = S
End If
End Function
