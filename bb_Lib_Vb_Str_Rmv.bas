Attribute VB_Name = "bb_Lib_Vb_Str_Rmv"
Option Compare Database
Option Explicit
Function RmvDblSpc$(S)
Dim O$: O = S
While HasSubStr(O, "  ")
    O = Replace(O, "  ", " ")
Wend
RmvDblSpc = O
End Function
Function RmvLasNChr$(S, Optional N% = 1)
RmvLasNChr = Left(S, Len(S) - 1)
End Function
Function RmvFstNChr(S, Optional N% = 1)
RmvFstNChr = Mid(S, N + 1)
End Function
Function RmvPFx$(S, Pfx)
Dim L%: L = Len(Pfx)
If Left(S, L) = Pfx Then
    RmvPFx = Mid(S, L + 1)
Else
    RmvPFx = S
End If
End Function
