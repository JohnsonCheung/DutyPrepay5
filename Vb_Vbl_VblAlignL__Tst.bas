Attribute VB_Name = "Vb_Vbl_VblAlignL__Tst"
Option Compare Database
Option Explicit
Private Type TstDta
    Vbl As String
    W As Integer
    FstVblNSpc As Integer
    RstVblNSpc As Integer
    Exp As String
    ThowMsg As String
End Type

Private Function Act$(A As TstDta)
With A
Act = VblAlignL(.Vbl, .W, .FstVblNSpc, .RstVblNSpc)
End With
End Function

Private Function ActTMOStr(A As TstDta) As ThowMsgOrStr
On Error GoTo X
ActTMOStr = TMOStrSomStr(Act(A))
Exit Function
X:
ActTMOStr = TMOStrThowMsg(Err.Description)
End Function

Private Function TstDta0() As TstDta
With TstDta0
    .Vbl = ""
    .ThowMsg = "Fun(VblAlignL) Prm(Vbl) has Er(Cannot be Blank)"
End With
End Function

Private Function TstDta1() As TstDta
With TstDta1
    .Vbl = "AAA|B    B|C"
    .FstVblNSpc = -1
    .ThowMsg = "Fun(Assert_FstVblNSpc_RstVblNSpc) Prm(FstVblNSpc) has Er(Cannot be negative)"
End With
End Function

Private Function TstDta2() As TstDta
With TstDta2
    .Vbl = "AAA|B    B|C"
    .RstVblNSpc = -1
    .ThowMsg = "Fun(Assert_FstVblNSpc_RstVblNSpc) Prm(RstVblNSpc) has Er(Cannot be negative)"
End With
End Function

Private Function TstDta3() As TstDta
With TstDta3
    .Vbl = "AAA|B    B|C"
    .FstVblNSpc = 4
    .RstVblNSpc = 6
    .W = 15
    .Exp = "    AAA        |      B    B   |      C        "
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

Private Sub TstDtaPush(O() As TstDta, M As TstDta)
Dim N&: N = TstDtaSz(O)
ReDim Preserve O(N)
O(N) = M
End Sub

Private Function TstDtaSz&(A() As TstDta)
On Error Resume Next
TstDtaSz = UBound(A) + 1
End Function

Private Sub Tstr(A As TstDta)
Dim M As ThowMsgOrStr: M = ActTMOStr(A)
With A
    If .ThowMsg = "" Then
        If Not M.Som Then Stop
        AssertActEqExp M.Str, .Exp
    Else
        If M.Som Then Stop
        AssertActEqExp M.ThowMsg, .ThowMsg
    End If
End With
End Sub

Sub VblAlignL__Tst()
Dim Ay() As TstDta
    Ay = TstDtaAy
Dim J%
For J = 0 To UBound(Ay)
    If J = J Then
    Tstr Ay(J)
    End If
Next
End Sub
