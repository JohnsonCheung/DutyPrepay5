VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

Function ObjSelPrp(Obj, PrpNy$()) As Variant()
Dim U%
    U = UB(PrpNy)
Dim O()
    ReDim O(U)
    Dim J%
    For J = 0 To U
        O(J) = CallByName(Obj, PrpNy(J), VbGet)
    Next
ObjSelPrp = O
End Function


Private Sub ObjCollSel__Tst()
'ObjCollSel(DbT("Permit").DaoFlds, "Name Type").Brw
End Sub

Sub Tst()

ObjCollSel__Tst
End Sub
