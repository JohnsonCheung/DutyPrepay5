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
