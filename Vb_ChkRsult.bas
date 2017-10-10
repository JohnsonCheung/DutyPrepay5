Attribute VB_Name = "Vb_ChkRsult"
Option Compare Database
Option Explicit
Type ChkRslt
    LinIdx As Long
    Msg As String
End Type
Function ChkRsltUB&(A() As ChkRslt)
ChkRsltUB = ChkRsltSz(A) - 1
End Function
Function ChkRsltLy(A() As ChkRslt) As String()
Dim O$(), J%
Push O, "ChkRsltCnt=" & ChkRsltSz(A)
For J = 0 To ChkRsltUB(A)
Push O, ChkRsltToStr(A(J))
Next
ChkRsltLy = O
End Function
Function ChkRsltToStr$(A As ChkRslt)
ChkRsltToStr = AlignL(A.LinIdx, 4) & ": [" & A.Msg & "]"
End Function
Sub ChkRsltDmp(A() As ChkRslt)
AyDmp ChkRsltLy(A)
End Sub
Function ChkRsltSz&(A() As ChkRslt)
On Error Resume Next
ChkRsltSz = UBound(A) + 1
End Function

Sub ChkRsltPush(O() As ChkRslt, I As ChkRslt)
Dim N&: N = ChkRsltSz(O)
ReDim Preserve O(N)
O(N) = I
End Sub
Sub ChkRsltPushAy(O() As ChkRslt, A() As ChkRslt)
Dim J&
For J = 0 To ChkRsltUB(A)
    ChkRsltPush O, A(J)
Next
End Sub
Sub ChkRsltPut__Tst()
Dim Ay$()
Dim R() As ChkRslt
Dim Act$()
Dim Exp$()
    
'---
Erase R
Ay = SplitLvs("aaa bbbb ccc")
ChkRsltPush R, ChkRslt("a er", 0)
ChkRsltPush R, ChkRslt("c er", 2)
ChkRsltPush R, ChkRslt("end msg")
Act = ChkRsltPut(Ay, R)
Exp = Sy("aaa  ---(a er)", "bbbb", "ccc  ---(c er)", "---(end msg)")
Debug.Assert AyIsEq(Act, Exp)
    
'---
Erase R
Ay = SplitLvs("aaa bbbb ccc")
ChkRsltPush R, ChkRslt("a1 er", 0)
ChkRsltPush R, ChkRslt("a2 er", 0)
ChkRsltPush R, ChkRslt("c er", 2)
ChkRsltPush R, ChkRslt("end msg")
Act = ChkRsltPut(Ay, R)
Exp = Sy("aaa  ---(a1 er) ---(a2 er)", "bbbb", "ccc  ---(c er)", "---(end msg)")
Debug.Assert AyIsEq(Act, Exp)
End Sub
Function ChkRsltPut(Ay$(), R() As ChkRslt) As String()
Dim EndMsg$()
Dim R1() As ChkRslt
    Dim J%
    For J = 0 To ChkRsltUB(R)
        If R(J).LinIdx = -1 Then
            Push EndMsg, R(J).Msg
        Else
            ChkRsltPush R1, R(J)
        End If
    Next
Dim O$()
    O = Ay
    For J = 0 To UB(EndMsg)
        Push O, "---(" & EndMsg(J) & ")"
    Next
    
    If ChkRsltSz(R1) > 0 Then
        O = AyAlignL(O)
        For J = 0 To ChkRsltUB(R1)
            With R1(J)
                O(.LinIdx) = O(.LinIdx) & " ---(" & .Msg & ")"
            End With
        Next
    End If
ChkRsltPut = O
End Function
Function ChkRsltMkEndMsg(EndMsg$()) As ChkRslt()
Dim O() As ChkRslt, J%
For J = 0 To UB(EndMsg)
    ChkRsltPush O, ChkRslt(EndMsg(J))
Next
ChkRsltMkEndMsg = O
End Function
Function ChkRslt(Msg, Optional LinIdx% = -1) As ChkRslt
ChkRslt.LinIdx = LinIdx
ChkRslt.Msg = Msg
End Function

