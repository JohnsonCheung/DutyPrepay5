Attribute VB_Name = "Vb_Str_Split"
Option Explicit
Option Compare Database

Function SplitCrLf(S) As String()
SplitCrLf = Split(S, vbCrLf)
End Function

Function SplitLvs(Lvs) As String()
SplitLvs = Split(RmvDblSpc(Trim(Lvs)), " ")
End Function

Function SplitSpc(S) As String()
SplitSpc = Split(S, " ")
End Function

Function SplitVBar(S) As String()
SplitVBar = Split(S, "|")
End Function
