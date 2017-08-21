VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Coll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Public Coll As Object
Function Sel(PrpNy) As Drs
Dim Ay$()
    Ay = Ny(PrpNy).Ny
Dim Dry As New Dry
    Dim Obj
'    If Not IsEmptyColl(ObjColl) Then
'        For Each Obj In ObjColl
'            Dry.Push ObjSelPrp(Obj, Ay)
'        Next
'    End If
Set Sel = Nw.Drs(Ay, Dry)
End Function
