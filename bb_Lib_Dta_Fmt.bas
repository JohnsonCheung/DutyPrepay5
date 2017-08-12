Attribute VB_Name = "bb_Lib_Dta_Fmt"
Option Compare Database
Option Explicit

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
Function DrLin$(Dr, Wdt%())
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
        O(J) = DrLin__V(V, W)
        J = J + 1
    Next
DrLin = Quote(Join(O, " | "), "| * |")
End Function

Private Function DrLin__V$(V, W)
Dim O$
If IsArray(V) Then
    If AyIsEmpty(V) Then
        O = AlignL("", W)
    Else
        O = AlignL(FmtQQ("Ay?:", UB(V)) & V(0), W)
    End If
Else
    O = Replace(O, vbCrLf, "|")
    O = AlignL(V, W)
End If
DrLin__V = O
End Function
