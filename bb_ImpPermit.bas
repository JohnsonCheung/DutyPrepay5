Attribute VB_Name = "bb_ImpPermit"
Option Compare Database
Option Explicit

Sub ImpPermit()
Dim Fx
Dim FxAy$(): FxAy = Pth(PermitImpPth).FfnAy("*.xlsx")
If AyIsEmpty(FxAy) Then Exit Sub
Dim M As PermImpFx
For Each Fx In FxAy
    Set M = PermImpFx(Fx)
Next
End Sub
