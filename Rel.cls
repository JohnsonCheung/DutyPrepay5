VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Rel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Private A_() As RelItm
Private Function ZZ() As Rel
Set ZZ = New Rel
End Function

Private Sub Srt__Tst()
Dim Act As Rel
With ZZ
    .Add RelItmLvs("A", "B C D")
    .Add RelItmLvs("B", "E F G")
    .Add RelItmLvs("C", "D D F")
    .Add RelItmLvs("D", "B C D")
    .Add RelItmLvs("E", "B C D")
    Set Act = Srt
End With
End Sub
Sub Add(Itm As RelItm)
If Has(Itm.Nm) Then Stop
Push Itm
End Sub
Function Has(Nm$) As Boolean
Dim J&
For J = 0 To UB
    If A_(J).Nm = Nm Then Has = True: Exit Function
Next
End Function
Property Get ItmAy() As RelItm()
ItmAy = A_
End Property
Property Let ItmAy(Ay() As RelItm)
A_ = Ay
End Property
Function Cpy() As Rel
Dim O As New Rel
O.ItmAy = A_
Set Cpy = O
End Function
Function Itm(Idx&) As RelItm
Itm = A_(Idx)
End Function
Function Srt() As Rel
Dim O As New Rel, Ay() As RelItm
Dim M As Rel: Set M = Cpy
Dim IdxAy&(), J&
IdxAy = M.ChdIdxAy
While Not AyIsEmpty(IdxAy)
    For J = 0 To bb_Lib_Vb_Ay.UB(IdxAy)
        O.Push M.Itm(J)
        M.RmvAt IdxAy(J)
    Next
    IdxAy = M.ChdIdxAy
Wend
If Not M.IsEmpty Then Stop
Set Srt = O
End Function
Property Get IsEmpty() As Boolean
IsEmpty = Sz = 0
End Property
Sub PushAy(Ay() As RelItm)
Dim J&
For J = 0 To UB
    Push Ay(J)
Next
End Sub
Function Nm$(Idx&)
Nm = A_(Idx).Nm
End Function
Function ChdAy() As String()
Dim J&, O$()
For J = 0 To UB
    If IdxIsChd(J) Then bb_Lib_Vb_Ay.Push O, Nm(J)
Next
ChdAy = O
End Function
Function ChdIdxAy() As Long()
Dim J&, O&()
For J = 0 To UB
    If IdxIsChd(J) Then bb_Lib_Vb_Ay.Push O, J
Next
ChdIdxAy = O
End Function
Function LeafIdxAy() As String()
Dim J&, O$()
For J = 0 To UB
    If IdxIsLeaf(J) Then bb_Lib_Vb_Ay.Push O, Nm(J)
Next
LeafIdxAy = O
End Function
Function LeafAy() As String()
Dim J&, O$()
For J = 0 To UB
    If IdxIsLeaf(J) Then bb_Lib_Vb_Ay.Push O, Nm(J)
Next
LeafAy = O
End Function
Function Idx&(Nm$)
Dim J&
For J = 0 To UB
    If A_(J).Nm = Nm Then Idx = J: Exit Function
Next
Idx = -1
End Function

Function IsLeaf(Nm$) As Boolean
IsLeaf = IdxIsLeaf(Idx(Nm))
End Function
Function IdxHasChd(Idx&) As Boolean
IdxHasChd = Not AyIsEmpty(A_(Idx).Chd)
End Function
Function IdxIsLeaf(I&) As Boolean
If Not IdxHasChd(I) Then IdxIsLeaf = True: Exit Function
Dim Chd, II&
For Each Chd In A_(I).Chd
    II = Idx(CStr(Chd))
    If II >= 0 Then
        If IdxHasChd(II) Then Exit Function
    End If
Next
IdxIsLeaf = True
End Function
Function IdxIsChd(I&) As Boolean
Dim Nm$: Nm = Me.Nm(I)
Dim J&
For J = 0 To UB
    If AyHas(Chd(J), Nm) Then Exit Function
Next
IdxIsChd = True
End Function
Function Chd(Idx&) As String()
Chd = A_(Idx).Chd
End Function
Function Sz&()
On Error Resume Next
Sz = UBound(A_) + 1
End Function
Function UB&()
UB = Sz - 1
End Function
Sub RmvItm(Nm$)
RmvAt Idx(Nm)
End Sub
Sub RmvAt(Idx&)
Dim J&
For J = Idx + 1 To UB
    A_(J - 1) = A_(J)
Next
DecSz
End Sub
Sub DecSz()
If IsEmpty Then Erase A_: Exit Sub
ReDim Preserve A_(UB - 1)
End Sub
Function LasEle() As RelItm
LasEle = A_(UB)
End Function
Sub Push(Itm As RelItm)
Dim N&: N = Sz
ReDim Preserve A_(N)
A_(N) = Itm
End Sub
Function Pop() As RelItm
Pop = LasEle
DecSz
End Function


