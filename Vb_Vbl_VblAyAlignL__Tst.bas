Attribute VB_Name = "Vb_Vbl_VblAyAlignL__Tst"
Option Compare Database
Option Explicit
Private Type TstDta
    VblAy As Variant
    FstVblNSpc As Integer
    RstVblNSpc As Integer
    Exp() As String
    ThowMsg As String
End Type

Private Function Act(A As TstDta) As String()
With A
Act = VblAyAlignL(.VblAy, .FstVblNSpc, .RstVblNSpc)
End With
End Function

Private Function ActTMOSy(A As TstDta) As ThowMsgOrSy
On Error GoTo X
ActTMOSy = TMOSySomSy(Act(A))
Exit Function
X:
ActTMOSy = TMOSySomThowMsg(Err.Description)
End Function

Private Function TstDta0() As TstDta
With TstDta0
    .VblAy = Sy()
    .ThowMsg = "Function-VblAyAlignL-Prm-VblAy-Error: Empty Ay"
End With
End Function

Private Function TstDta1() As TstDta
With TstDta1
    .VblAy = Sy("AAA|B    B|C", "AA")
    .FstVblNSpc = -1
    .ThowMsg = "Function-VblIndent-Prm-FstVblNSpc-Error: Cannot be negative"
End With
End Function

Private Function TstDta2() As TstDta
With TstDta2
    .VblAy = Sy("AAA|B    B|C", "AA")
    .RstVblNSpc = -1
    .ThowMsg = "Function-VblIndent-Prm-RstVblNSpc-Error: Cannot be negative"
End With
End Function

Private Function TstDta3() As TstDta
With TstDta3
    .VblAy = Sy("AAA|B    B|C", "AA")
    .FstVblNSpc = 4
    .RstVblNSpc = 6
    .Exp = Sy("    AAA|      B    B|    C", "    AA")
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
Dim M As ThowMsgOrSy: M = ActTMOSy(A)
With A
    If .ThowMsg = "" Then
        If Not M.Som Then Stop
        AssertActEqExp M.Sy, .Exp
    Else
        If M.Som Then Stop
        AssertActEqExp M.ThowMsg, .ThowMsg
    End If
End With
End Sub

Sub VblAyAlignL__Tst()
Dim Ay() As TstDta
    Ay = TstDtaAy
Dim J%
For J = 3 To 3 ' UBound(Ay)
    Tstr Ay(J)
Next
End Sub
