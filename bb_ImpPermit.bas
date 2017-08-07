Attribute VB_Name = "bb_ImpPermit"
Option Compare Database
Option Explicit
Sub ImpPermit()
Dim Fx
Dim mFxAy$(): mFxAy = FxAy()
If IsEmptyAy(mFxAy) Then Exit Sub
Dim M As Permit
For Each Fx In mFxAy
    Set M = New Permit
    M.Impport Fx
Next
End Sub

Private Function FxAy() As String()
FxAy = PthFfnAy(PermitImpPth, "*.xlsx")
End Function

Private Sub ImpPermit__Tst()
Dim M As Permit
Set M = New Permit
M.TstEr
'M.TstOk
End Sub
