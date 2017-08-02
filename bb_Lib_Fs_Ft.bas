Attribute VB_Name = "bb_Lib_Fs_Ft"
Option Compare Database
Option Explicit
Function OpnApp%(Ft)
Dim O%: O = FreeFile(1)
Open Ft For Append As #O
OpnApp = O
End Function
Function OpnOup%(Ft)
Dim O%: O = FreeFile(1)
Open Ft For Output As #O
OpnOup = O
End Function
Function OpnInp%(Ft)
Dim O%: O = FreeFile(1)
Open Ft For Input As #O
OpnInp = O
End Function
Sub BrwFt(Ft)
Shell "NotePad """ & Ft & """", vbMaximizedFocus
End Sub
