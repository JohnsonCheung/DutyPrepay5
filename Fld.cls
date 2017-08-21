VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Fld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Public Fld As Dao.Field
Function Init(A As Dao.Field) As Fld
Set Fld = A
Set Init = Me
End Function

Function Des$()
Des = Prps.Des
End Function

Function Prps() As Prps
Set Prps = Nw.Prps(Fld.Properties)
End Function

Function InfDr(Pk$()) As Variant()
With Fld
    InfDr = Array(.Name, IIf(Ay(Pk).Has(.Name), "*", ""), DaoTyStr(.Type), .Size, .DefaultValue, .Required, Des)
End With
End Function

Function InfFny() As String()
InfFny = SplitSpc("Fld Pk Ty Sz Dft Req Des")
End Function

Function InsSqlValPhraseTp$()
InsSqlValPhraseTp = SimTyInsSqlValPhraseTp(DaoTySimTy(Fld.Type))
End Function

