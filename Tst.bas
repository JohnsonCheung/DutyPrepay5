Attribute VB_Name = "Tst"
Option Explicit
Option Compare Database

Function TstResPth$()
TstResPth = PjSrcPth & "TstRes\"
End Function

Sub TstResPthBrw()
PthBrw TstResPth
End Sub
