Attribute VB_Name = "Vb_LnoCnt"
Option Explicit
Option Compare Database
Type LnoCnt
    Lno As Long
    Cnt As Long
End Type

Function LnoCnt(Lno&, Cnt&) As LnoCnt
LnoCnt.Lno = Lno
LnoCnt.Cnt = Cnt
End Function

Function LnoCntAyIsEmpty(A() As LnoCnt) As Boolean
LnoCntAyIsEmpty = LnoCntSz(A) = 0
End Function

Sub LnoCntDmp(A As LnoCnt)
Debug.Print LnoCntToStr(A)
End Sub

Sub LnoCntPush(O() As LnoCnt, M As LnoCnt)
Dim N&: N = LnoCntSz(O)
ReDim Preserve O(N)
O(N) = M
End Sub

Function LnoCntStr$(A As LnoCnt)
LnoCntStr = FmtQQ("Lno(?) Cnt(?)", A.Lno, A.Cnt)
End Function

Function LnoCntSz&(A() As LnoCnt)
On Error Resume Next
LnoCntSz = UBound(A) + 1
End Function

Function LnoCntToStr$(A As LnoCnt)
With A
    LnoCntToStr = FmtQQ("Lno(?) Cnt(?)", .Lno, .Cnt)
End With
End Function

Function LnoCntUB&(A() As LnoCnt)
LnoCntUB = LnoCntSz(A) - 1
End Function
