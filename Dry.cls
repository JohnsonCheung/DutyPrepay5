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
Dim Dry_()
Property Get Dry() As Variant()
Dry = Dry_
End Property
Sub Push(Dr)
If Not IsArray(Dr) Then Stop
bb_Lib_Vb_Ay.Push Dry_, Dr
End Sub
Function NRow&()
On Error Resume Next
NRow = UBound(Dry_) + 1
End Function
Function URow&()
URow = NRow - 1
End Function
Sub PushAy(DrAy)
Dim Dr
If AyIsEmpty(DrAy) Then Exit Sub
For Each Dr In DrAy
    Push Dr
Next
End Sub
Property Get Dr(Idx&) As Variant()
Dr = Dry_(Idx)
End Property
Property Let Dr(Idx&, V())
Dry_(Idx) = V
End Property

Function IsEmpty() As Boolean
IsEmpty = NRow = 0
End Function
Function Col(Optional ColIdx% = 0) As Variant()
If IsEmpty Then Exit Function
Dim O(), Dr
For Each Dr In Dry_
    bb_Lib_Vb_Ay.Push O, Dr(ColIdx)
Next
Col = O
End Function

Sub Dmp()
Ly.Ay.Dmp
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
Set Sq = ccNew.Sq(O)
End Function

Function StrCol(Dry, Optional ColIdx% = 0) As String()
StrCol = AySy(Col(ColIdx))
End Function
Sub Brw()
Ly.Ay.Brw
End Sub
Function Ly(Optional MaxColWdt% = 100, Optional BrkColIdx% = -1) As Sy
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
Dim O$()
    Dim Dr
    Dim LasV, V
    bb_Lib_Vb_Ay.Push O, Hdr
    Dim IsBrk As Boolean
    If BrkColIdx >= 0 Then LasV = Ay(0)(BrkColIdx)
    For Each Dr In Dry
        IsBrk = False
            If BrkColIdx >= 0 Then
                V = Dr(BrkColIdx)
                If LasV <> V Then
                    IsBrk = True
                    LasV = V
                End If
            End If
        If IsBrk Then bb_Lib_Vb_Ay.Push O, Hdr
        bb_Lib_Vb_Ay.Push O, DrLin(Dr, W)
    Next
    bb_Lib_Vb_Ay.Push O, Hdr
Set Ly = Sy(O)
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

