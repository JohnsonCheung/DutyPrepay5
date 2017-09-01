Attribute VB_Name = "Tst"
Option Compare Database
Option Explicit

Function TstResPth$()
TstResPth = PjSrcPth & "TstRes\"
End Function

Sub TstResPthBrw()
PthBrw TstResPth
End Sub
