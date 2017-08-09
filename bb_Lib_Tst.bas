Attribute VB_Name = "bb_Lib_Tst"
Option Compare Database
Option Explicit

Function TstResPth$()
TstResPth = PjSrcPth & "TstRes\"
End Function

Sub TstResPthBrw()
PthBrw TstResPth
End Sub
