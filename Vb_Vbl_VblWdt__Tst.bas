Attribute VB_Name = "Vb_Vbl_VblWdt__Tst"
Option Compare Database
Option Explicit
Private Type TstDta
    Vbl As String
    RstVblNSpc As Integer
    FstVblNSpc As Integer
    Exp As Integer
    ThowMsg As String
End Type

Sub TstrDmp(A As TstDta)
End Sub

Private Function Act%(A As TstDta)
With A
    Act = VblWdt(.Vbl, .FstVblNSpc, .RstVblNSpc)
End With
End Function

Private Function ActTMOInt(A As TstDta) As ThowMsgOrInt
On Error GoTo X
ActTMOInt = TMOIntSomInt(Act(A))
Exit Function
X:
ActTMOInt = TMOIntThowMsg(Err.Description)
End Function

Private Function TstDta0() As TstDta
With TstDta0
    .Vbl = "aaa|b|c     d"
    .RstVblNSpc = -2
    .FstVblNSpc = 4
    .ThowMsg = "Fun(Assert_FstVblNSpc_RstVblNSpc) Prm(RstVblNSpc) has Er(Cannot be negative)"
End With
End Function

Private Function TstDta1() As TstDta
With TstDta1
    .Vbl = "aaa|b|c     d"
    .RstVblNSpc = 0
    .FstVblNSpc = -4
    .ThowMsg = "Fun(Assert_FstVblNSpc_RstVblNSpc) Prm(FstVblNSpc) has Er(Cannot be negative)"
End With
End Function

Private Function TstDta2() As TstDta
With TstDta1
    .Vbl = "aaa|b|c     d"
    .FstVblNSpc = 1
    .RstVblNSpc = 2
    .Exp = 9
End With
End Function

Private Function TstDta3() As TstDta
With TstDta3
    .Vbl = "aaa|b|c     d"
    .FstVblNSpc = 2
    .RstVblNSpc = 4
    .Exp = 11
End With

End Function

Private Function TstDta4() As TstDta
With TstDta4
    .Vbl = "aaa|b|c     d"
    .FstVblNSpc = 4
    .RstVblNSpc = 2
    .Exp = 9
End With

End Function

Private Function TstDtaAy() As TstDta()
Dim O() As TstDta
TstDtaPush O, TstDta0
TstDtaPush O, TstDta1
TstDtaPush O, TstDta2
TstDtaPush O, TstDta3
TstDtaAy = O
End Function

Private Sub TstDtaPush(O() As TstDta, I As TstDta)
Dim N&: N = TstDtaSz(O)
ReDim Preserve O(N)
O(N) = I
End Sub

Private Function TstDtaSz&(A() As TstDta)
On Error Resume Next
TstDtaSz = UBound(A) + 1
End Function

Private Sub Tstr(A As TstDta, Dmp As Boolean)
Dim Act As ThowMsgOrInt
Act = ActTMOInt(A)
With A
    If .ThowMsg = "" Then
        AssertActEqExp Act.Int, A.Exp
    Else
        AssertActEqExp Act.ThowMsg, A.ThowMsg
    End If
End With
If Dmp Then TstrDmp A
End Sub

Sub VblWdt__Tst(Optional Dmp As Boolean)
Dim Ay() As TstDta
    Ay = TstDtaAy
    Dim J%
    For J = 0 To UBound(Ay)
        Tstr Ay(J), Dmp
    Next
End Sub
