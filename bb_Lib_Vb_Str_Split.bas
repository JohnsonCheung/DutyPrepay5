Attribute VB_Name = "bb_Lib_Vb_Str_Split"
Option Compare Database
Option Explicit

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
