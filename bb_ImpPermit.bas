Attribute VB_Name = "bb_ImpPermit"
Option Compare Database
Option Explicit

Sub ImpPermit()
Dim mFxAy$(): mFxAy = FxAy
If IsEmptyAy(mFxAy) Then Exit Sub
Dim M As Permit, Fx
For Each Fx In mFxAy
    Set M = New Permit
    M.Import (Fx)
Next
End Sub


Private Property Get FxAy() As String()
FxAy = PthFfnAy(PermitImpPth, "*.xlsx")
End Property

Sub AA()
ImpPermit
End Sub
