VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Private Dry_
Property Get Dry()
Dry = Dry_
End Property
Function Init(Dry) As Dry
AssertIsAy Dry
Dry_ = Dry
Set Init = Me
End Function
Function AddDr(Dr) As Dry
AsrtIsAy Dr
Vb.Push Dry_, Dr
Set AddDr = Me
End Function

Function NRow&()
On Error Resume Next
NRow = UBound(Dry_) + 1
End Function
Sub PutAt(AtCell As Range)
Sq.PutAt AtCell
End Sub
Function URow&()
URow = NRow - 1
End Function
Sub AddDrAy(DrAy)
Dim Dr
If AyIsEmpty(DrAy) Then Exit Sub
For Each Dr In DrAy
    AddDr Dr
Next
End Sub
Sub Tst()

End Sub
Function CsvQQStr$()
If IsEmpty Then Exit Function
CsvQQStr = DrX(0).CsvQQStr
End Function
Property Get Dr(RIdx&)
Dr = Dry_(RIdx)
End Property
Property Let Dr(RIdx&, V)
AsrtIsAy V
Dry_(RIdx) = V
End Property
Function DrX(RIdx&) As Ay
Set DrX = Nw.Ay(Dr(RIdx))
End Function
Function IsEmpty() As Boolean
IsEmpty = NRow = 0
End Function
Function Col(Optional ColIdx% = 0) As Variant()
If IsEmpty Then Exit Function
Dim O(), Dr
For Each Dr In Dry_
    Vb.Push O, Dr(ColIdx)
Next
Col = O
End Function
Function LyX(Optional MaxColWdt% = 100) As Ay
Set LyX = Nw.Ay(Ly(MaxColWdt))
End Function
Sub Dmp()
LyX.Dmp
End Sub

Function NCol%()
Dim Dr, O%, M%
For Each Dr In Dry
    M = Sz(Dr)
    If M > O Then O = M
Next
NCol = O
End Function

Function Sq(Optional NCol% = 0) As Sq
If AyIsEmpty(Dry) Then Exit Function
Dim NRow&
    If NCol = 0 Then NCol = Me.NCol
        NRow = Sz(Dry)
Dim O()
    ReDim O(1 To NRow, 1 To NCol)
Dim C%, R&, Dr
    R = 0
    For Each Dr In Dry
        R = R + 1
        For C = 0 To UB(Dr)
            O(R, C + 1) = Dr(C)
        Next
    Next
Set Sq = Nw.Sq(O)
End Function

Function StrCol(Dry, Optional ColIdx% = 0) As String()
StrCol = AySy(Col(ColIdx))
End Function
Sub Brw()
LyX.Brw
End Sub
Function Brk(BrkColIdx%, Optional MaxColWdt% = 100) As BrkLin
Set Brk = Nw.BrkLinByDryLy(Ly(MaxColWdt), BrkColIdx)
End Function
Function Ly(Optional MaxColWdt% = 100) As String()
If IsEmpty Then Exit Function
Dim W%(): W = WdtAy(MaxColWdt)
If AyIsEmpty(W) Then Exit Function
Dim HdrAy$()
    ReDim HdrAy(UB(W))
    Dim J%
    For J = 0 To UB(W)
        HdrAy(J) = String(W(J), "-")
    Next
Dim Hdr$: Hdr = Quote(Join(HdrAy, "-|-"), "|-*-|")
Dim O$(), Dr
    Push O, Hdr
    For Each Dr In Dry
        Vb.Push O, DrLin(Dr, W)
    Next
    Push O, Hdr
Ly = O
End Function

Function WdtAy(Optional MaxColWdt% = 100) As Integer()
If IsEmpty Then Exit Function
Dim O%()
    Dim Dr, UDr%, U%, V, L%, J%
    U = -1
    For Each Dr In Dry
        UDr = UB(Dr)
        If UDr > U Then ReDim Preserve O(UDr): U = UDr
        If AyIsEmpty(Dr) Then GoTo Nxt
        For J = 0 To UDr
            V = Dr(J)
            L = VarLen(V)
            
            If L > O(J) Then O(J) = L
        Next
Nxt:
    Next
Dim M%
M = MaxColWdt
For J = 0 To UB(O)
    If O(J) > M Then O(J) = M
Next
WdtAy = O
End Function

