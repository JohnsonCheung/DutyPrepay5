Attribute VB_Name = "cc_Rel"
Option Compare Database
Option Explicit
Public Type RelItm
    Nm As String
    Chd() As String
    Dta As Variant
End Type

Function RelItm(Nm$, Chd$(), Optional Dta) As RelItm
Dim O As RelItm
With O
    .Nm = Nm
    .Chd = Chd
    Asg Dta, .Dta
End With
RelItm = O
End Function
Function RelItmLvs(Nm$, ChdLvs$, Optional Dta) As RelItm
RelItmLvs = RelItm(Nm, SplitLvs(ChdLvs), Dta)
End Function

