Attribute VB_Name = "Dta_Fmt"
Option Explicit
Option Compare Database

Function DrsLy(a As Drs, Optional MaxColWdt& = 100, Optional BrkColNm$) As String()
If AyIsEmpty(a.Fny) Then Exit Function
Dim Drs As Drs: Drs = DrsAddRowIdxCol(a)
Dim Dry(): Dry = Drs.Dry
Push Dry, Drs.Fny
Dim Ay$(): Ay = DryLy(Dry, MaxColWdt)
Dim Lin$: Lin = Pop(Ay)
Dim Hdr$: Hdr = Pop(Ay)
Dim O$()
    PushAy O, Array(Lin, Hdr)
    PushAy O, Ay
    Push O, Lin
If BrkColNm <> "" Then O = DrsLyInsBrkLin(O, BrkColNm)
DrsLy = O
End Function

Function DrsLyInsBrkLin(TblLy$(), ColNm$) As String()
Dim Hdr$: Hdr = TblLy(1)
Dim Fny$():
    Fny = SplitVBar(Hdr)
    AyRmvFstEle Fny
    AyRmvLasEle Fny
    AyTrim Fny
Dim Idx%
    Idx = AyIdx(Fny, ColNm)
Dim DryLy$()
    DryLy = TblLy
    AyRmvEleAtCnt DryLy, 0, 2
Dim O$()
    Push O, TblLy(0)
    Push O, TblLy(1)
    PushAy O, DryLyInsBrkLin(DryLy, Idx)
DrsLyInsBrkLin = O
End Function

Sub DryAddBrkDr(ODry)
Dim W%(): W = DryWdtAy(ODry)
Dim O(), I
For Each I In W
    Push O, String(I, "-")
Next
Push ODry, O
End Sub

Function DryLy(Dry, Optional MaxColWdt& = 100, Optional BrkColIdx% = -1) As String()
If IsEmpty(Dry) Then Exit Function
Dim W%(): W = DryWdtAy(Dry, MaxColWdt)
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
    Push O, Hdr
    Dim IsBrk As Boolean
    If BrkColIdx >= 0 Then LasV = Dry(0)(BrkColIdx)
    For Each Dr In Dry
        IsBrk = False
            If BrkColIdx >= 0 Then
                V = Dr(BrkColIdx)
                If LasV <> V Then
                    IsBrk = True
                    LasV = V
                End If
            End If
        If IsBrk Then Push O, Hdr
        Push O, DrLin(Dr, W)
    Next
    Push O, Hdr
DryLy = O
End Function

Function DryLyInsBrkLin(DryLy$(), ColIdx%) As String()
If Sz(DryLy) = 2 Then DryLyInsBrkLin = DryLy: Exit Function
Dim Hdr$: Hdr = DryLy(0)
Dim Fm&, L%
    Dim N%: N = ColIdx + 1
    Dim P1&, P2&
    P1 = InstrN(Hdr, "|", N)
    P2 = InStr(P1 + 1, Hdr, "|")
    Fm = P1 + 1
    L = P2 - P1 - 1
Dim O$()
    Push O, DryLy(0)
    Dim LasV$: LasV = Mid(DryLy(1), Fm, L)
    Dim J&
    Dim V$
    For J = 1 To UB(DryLy) - 1
        V = Mid(DryLy(J), Fm, L)
        If LasV <> V Then
            Push O, Hdr
            LasV = V
        End If
        Push O, DryLy(J)
    Next
    Push O, AyLasEle(DryLy)
DryLyInsBrkLin = O
End Function

Function DtLy(Dt As Dt, Optional MaxColWdt& = 100, Optional BrkColNm$) As String()
Dim Rs As Drs
    Rs.Fny = Dt.Fny
    Rs.Dry = Dt.Dry
Dim O$()
    Push O, "*Tbl " & Dt.DtNm
    PushAy O, DrsLy(Rs, MaxColWdt, BrkColNm)
DtLy = O
End Function

Function VarLinStr$(V)
Dim O$
If IsEmpty(V) Then Exit Function
If IsObject(V) Then
    VarLinStr = "[" & TypeName(V) & "]"
    Exit Function
End If
If IsArray(V) Then
    If AyIsEmpty(V) Then Exit Function
    VarLinStr = "Ay?:" & UB(V) & V(0)
    Exit Function
End If
If InStr(V, vbCrLf) > 0 Then
    VarLinStr = Brk(V, vbCrLf).S1 & "|.."
    Exit Function
End If
VarLinStr = V
End Function

Private Function DrLin$(Dr, Wdt%())
Dim UDr%
    UDr = UB(Dr)
Dim O$()
    Dim U1%: U1 = UB(Wdt)
    ReDim O(U1)
    Dim W, V
    Dim J%
    J = 0
    For Each W In Wdt
        If UDr >= J Then V = Dr(J) Else V = ""
        O(J) = VarLinStrW(V, W)
        J = J + 1
    Next
DrLin = Quote(Join(O, " | "), "| * |")
End Function

Private Function DryWdtAy(Dry, Optional MaxColWdt& = 100) As Integer()
If AyIsEmpty(Dry) Then Exit Function
Dim O%()
    Dim Dr, UDr%, U%, V, L%, J%
    U = -1
    For Each Dr In Dry
        UDr = UB(Dr)
        If UDr > U Then ReDim Preserve O(UDr): U = UDr
        If AyIsEmpty(Dr) Then GoTo Nxt
        For J = 0 To UDr
            V = Dr(J)
            L = VarLinStrLen(V)
            
            If L > O(J) Then O(J) = L
        Next
Nxt:
    Next
Dim M%
M = MaxColWdt
For J = 0 To UB(O)
    If O(J) > M Then O(J) = M
Next
DryWdtAy = O
End Function

Private Function VarLinStrLen%(V)
VarLinStrLen = Len(VarLinStr(V))
End Function

Private Function VarLinStrW$(V, W)
VarLinStrW = AlignL(VarLinStr(V), W)
End Function

Private Sub DrsLyInsBrkLin__Tst()
Dim TblLy$()
Dim Act$()
Dim Exp$()
TblLy = FtLy(TstResPth & "DrsLyInsBrkLin.txt")
Act = DrsLyInsBrkLin(TblLy, "Tbl")
Exp = FtLy(TstResPth & "DrsLyInsBrkLin_Exp.txt")
AyAssertEq Exp, Act
End Sub

Sub Tst()
DrsLyInsBrkLin__Tst
End Sub
