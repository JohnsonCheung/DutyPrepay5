Attribute VB_Name = "Vb_Vbl"
'Vbl is Vbl.  It is a string without VbCr and VbLf.  It uses | as VbCrLf.  It can be converted to Lines.
Option Compare Database
Option Explicit

Function VblAlignL$(Vbl, W%, Optional FstVblNSpc%, Optional RstVblNSpc%)
Const CSub$ = "VblAlignL"
Assert_FstVblNSpc_RstVblNSpc FstVblNSpc, RstVblNSpc
VblAssertIsVdt Vbl
AssertNonEmpty Vbl
Dim FstSpc$, RstSpc$
    FstSpc = Space(FstVblNSpc)
    RstSpc = Space(RstVblNSpc)
Dim O$()
    Dim Ay$()
    Ay = SplitVBar(Vbl)
    Push O, AlignL(FstSpc & Ay(0), W, ErIfNotEnoughWdt:=True)
    Dim J%
    For J = 1 To UB(Ay)
        Push O, AlignL(RstSpc & Ay(J), W, ErIfNotEnoughWdt:=True)
    Next
VblAlignL = JnVBar(O)
End Function

Sub VblAssertIsVdt(Vbl)
Const CSub$ = "VbAssertIsVdt"
AssertIsStr Vbl
AssertNotHasSubStr Vbl, vbCr
AssertNotHasSubStr Vbl, vbLf
End Sub

Function VblAyAlignL(VblAy, Optional FstVblNSpc%, Optional RstVblNSpc%) As String()
Const CSub$ = "VblAyAlignL"
If AyIsEmpty(VblAy) Then Er CSub, "[VblAy] cannot be empty"
Dim WdtAy%()
    Dim Vbl
    For Each Vbl In VblAy
        Push WdtAy, VblWdt(Vbl, FstVblNSpc, RstVblNSpc)
    Next
Dim W%
    W = AyMax(WdtAy)
Dim O$()
    For Each Vbl In VblAy
        Push O, VblAlignL(Vbl, W, FstVblNSpc, RstVblNSpc)
    Next
VblAyAlignL = O
End Function

Sub VblAyAssertIsVdt(V)
AssertIsAy V
If AyIsEmpty(V) Then Exit Sub
Dim Vbl
For Each Vbl In V
    VblAssertIsVdt Vbl
Next
End Sub

Function VblByLines$(Lines)
If HasSubStr(Lines, "|") Then Stop
VblByLines = Replace(Lines, vbCrLf, "|")
End Function

Function VblLasLin$(Vbl)
VblLasLin = AyLasEle(SplitVBar(Vbl))
End Function

Function VblLines$(Vbl)
VblLines = Replace(Vbl, "|", vbCrLf)
End Function

Function VblLy(Vbl) As String()
VblLy = SplitVBar(Vbl)
End Function

Function VblWdt%(Vbl, Optional FstVblNSpc%, Optional RstVblNSpc%)
Assert_FstVblNSpc_RstVblNSpc FstVblNSpc, RstVblNSpc
VblAssertIsVdt Vbl
Dim WdtAy%()
Dim L, J%
Dim Ay$(): Ay = VblLy(Vbl)
For J = 0 To UB(Ay)
    If J = 0 Then
        Push WdtAy, FstVblNSpc + Len(Ay(J))
    Else
        Push WdtAy, RstVblNSpc + Len(Ay(J))
    End If
Next
VblWdt = AyMax(WdtAy)
End Function

Private Sub Assert_FstVblNSpc_RstVblNSpc(FstVblNSpc%, RstVblNSpc%)
Const CSub$ = "Assert_FstVblNSpc_RstVblNSpc"
If FstVblNSpc < 0 Then Er CSub, "{FstVblNSpc} Cannot be negative", FstVblNSpc
If RstVblNSpc < 0 Then Er CSub, "{RstVblNSpc} Cannot be negative", RstVblNSpc
End Sub
