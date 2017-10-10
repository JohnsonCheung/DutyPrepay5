Attribute VB_Name = "Sqp_SqpGp__Tst"
Option Compare Database
Option Explicit
Private Type TstDta
    ExprVblAy As Variant
    Exp As String
    Thow As Boolean
End Type
Const E1$ = "Expr1"
Const E2$ = "Expr2-Line1|Expr2-Line2"
Const E3$ = "Expr3-Line1|Expr3-Line2|Expr3-Line3"

Private Function Act$(ExprVblAy, Thow As Boolean)
If Thow Then
    On Error GoTo X
    Act = SqpGp(ExprVblAy)
    Stop ' Calling SqpGp should thow, but it does not
X:  Exit Function
End If
Act = SqpGp(ExprVblAy)
End Function

Private Function TstDta0() As TstDta
With TstDta0
    .ExprVblAy = Array(E1, E2, E3)
    .Thow = False
    .Exp = ""
End With
End Function

Private Function TstDta1() As TstDta
With TstDta1
    .ExprVblAy = Array()
    .Exp = ""
    .Thow = True
End With
End Function

Private Function TstDta2() As TstDta

End Function

Private Function TstDta3() As TstDta

End Function

Private Function TstDta4() As TstDta

End Function

Private Function TstDta5() As TstDta

End Function

Private Function TstDtaAy() As TstDta()
Dim J%, O() As TstDta
TstDtaPush O, TstDta0
TstDtaPush O, TstDta1
TstDtaPush O, TstDta2
TstDtaPush O, TstDta3
TstDtaPush O, TstDta4
TstDtaPush O, TstDta5
TstDtaAy = O
End Function

Private Sub TstDtaPush(O() As TstDta, I As TstDta)
Dim N&: N = TstDtaSz(O)
ReDim Preserve O(N)
O(N) = I
End Sub

Private Function TstDtaSz&(O() As TstDta)
On Error Resume Next
TstDtaSz = UBound(O) + 1
End Function

Private Sub Tstr(A As TstDta)
Dim mAct$
    mAct = Act(A.ExprVblAy, A.Thow)
If A.Thow Then Exit Sub
With A
    AssertActEqExp mAct, .Exp
End With
End Sub

Private Sub SqpGp__Tst()
Dim Ay() As TstDta
Ay = TstDtaAy
Dim J%
For J = 0 To 0 ' UBound(Ay)
    Tstr Ay(J)
Next
End Sub
