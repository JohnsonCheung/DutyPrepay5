Attribute VB_Name = "bb_ImpPermit"
Option Compare Database
Option Explicit
Sub ImpPermit()
Dim FxAy$(): FxAy = A_FxAy
If IsEmptyAy(FxAy) Then Exit Sub
Dim Fx
For Each Fx In FxAy
    NewPermit(Fx).Import
Next
End Sub
Function NewPermit(Fx) As Permit
Dim O As New Permit
O.Fx = Fx
Set NewPermit = O
End Function

Private Property Get A_FxAy() As String()
A_FxAy = PthFfnAy(PermitImpPth, "*.xlsx")
End Property

Sub AA()
ImpPermit
End Sub
