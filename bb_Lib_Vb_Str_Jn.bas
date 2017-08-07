Attribute VB_Name = "bb_Lib_Vb_Str_Jn"
Option Compare Database

Function JnComma$(Ay)
JnComma = Join(Ay, ",")
End Function

Function JnCrLf(Ay)
JnCrLf = Join(Ay, vbCrLf)
End Function

Function JnSpc(Ay)
JnSpc = Join(Ay, " ")
End Function
