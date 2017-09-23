Attribute VB_Name = "Tst"
Option Explicit
Option Compare Database

Function TstResFdr$(Fdr$)
Dim O$
    O = TstResPth & Fdr & "\"
    PthEns O
TstResFdr = O
End Function

Sub TstResFdrBrw(Fdr$)
PthBrw TstResFdr(Fdr)
End Sub

Function TstResPth$()
Dim O$
    O = PjSrcPth & "TstRes\"
    PthEns O
TstResPth = O
End Function

Sub TstResPthBrw()
PthBrw TstResPth
End Sub
