Attribute VB_Name = "bb_Lib_Tst"
Option Compare Database
Option Explicit

Function TstResPth$()
TstResPth = CurPj.SrcPth & "TstRes\"
End Function

Sub TstResPthBrw()
PthBrw TstResPth
End Sub