Attribute VB_Name = "bb_Lib_Vb_Str_Split"
Option Compare Database
Option Explicit
Function SplitLvs(Lvs) As String()
SplitLvs = Split(RmvDblSpc(Trim(Lvs)), " ")
End Function
Function SplitSpc(S) As String()
SplitSpc = Split(S, " ")
End Function
