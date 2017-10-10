Attribute VB_Name = "Vb_FmTo"
Option Explicit
Option Compare Database

Type FmTo
    FmIdx As Long
    ToIdx As Long
End Type

Function EmptyFmTo() As FmTo
EmptyFmTo.FmIdx = -1
EmptyFmTo.ToIdx = -1
End Function

Function FmTo(FmIdx&, ToIdx&) As FmTo
FmTo.FmIdx = FmIdx
FmTo.ToIdx = ToIdx
End Function

Function FmToAyIsEmpty(A() As FmTo) As Boolean
FmToAyIsEmpty = FmToSz(A) = 0
End Function

Function FmToAyLnoCntAy(A() As FmTo) As LnoCnt()
If FmToAyIsEmpty(A) Then Exit Function
Dim U&, J&
U = FmToUB(A)
Dim O() As LnoCnt
    ReDim O(U)
For J = 0 To U
    O(J) = FmToLnoCnt(A(J))
Next
FmToAyLnoCntAy = O
End Function

Function FmToLnoCnt(A As FmTo) As LnoCnt
Dim Lno&, Cnt&
    Cnt = A.ToIdx - A.FmIdx + 1
    If Cnt < 0 Then Cnt = 0
    Lno = A.FmIdx + 1
With FmToLnoCnt
    .Cnt = Cnt
    .Lno = Lno
End With
End Function

Function FmToN&(A As FmTo)
With A
    FmToN = .ToIdx - .FmIdx + 1
End With
End Function

Sub FmToPush(O() As FmTo, M As FmTo)
Dim N&: N = FmToSz(O)
ReDim Preserve O(N)
O(N) = M
End Sub

Function FmToSz&(A() As FmTo)
On Error Resume Next
FmToSz = UBound(A) + 1
End Function

Function FmToUB&(A() As FmTo)
FmToUB = FmToSz(A) - 1
End Function
