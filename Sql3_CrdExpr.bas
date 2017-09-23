Attribute VB_Name = "Sql3_CrdExpr"
Option Compare Database
Option Explicit
Private CrdTyLvs$

Function CrdExpr$(CrdTyLvs_$)
CrdTyLvs = CrdTyLvs_
        Const CaseWhen$ = "Case When"
CrdExpr = CaseWhen & Gp & ElseN & EndN
End Function

Private Function CrdPfxAy(CrdTyId%) As String()
Dim Dry(): Dry = DrySel(SR_CrdPfxTyDry, 1, CrdTyId)
CrdPfxAy = DryStrCol(Dry, 0)
End Function

Private Function CrdTyAy() As Integer()
CrdTyAy = AyAsgInto(SplitLvs(CrdTyLvs), EmptyIntAy)
End Function

Private Function ElseN$()
ElseN = "|Else " & NGp + 1
End Function

Private Function EndN$()
EndN = "|" & StrDup(NGp, "End ")
End Function

Private Function Gp$()
Const ElseCaseWhen$ = "|Else Case When"
Gp = Join(GpAy, ElseCaseWhen)
End Function

Private Function GpAy() As String()
Dim O$(), J%, mCrdTyAy%()
mCrdTyAy = CrdTyAy
For J = 0 To UB(CrdTyAy)
    Push O, GpItm(mCrdTyAy(J))
Next
GpAy = O
End Function

Private Function GpItm$(CrdTyId%)
Dim Ay$(): Ay = SHMCodeLikAy(CrdTyId)
Const Sep$ = " OR"
GpItm = Join(Ay, Sep) & " THEN " & CrdTyId
End Function

Private Function NGp%()
NGp = Sz(CrdTyAy)
End Function

Private Function SHMCodeLik(Pfx)
SHMCodeLik = FmtQQ("|SHMCode Like '?%'", Pfx)
End Function

Private Function SHMCodeLikAy(CrdTyId%) As String()
Dim O$(), Pfx
Dim PfxAy$(): PfxAy = CrdPfxAy(CrdTyId)
For Each Pfx In PfxAy
    Push O, SHMCodeLik(Pfx)
Next
O = AyAlignL(O)
SHMCodeLikAy = O
End Function

Private Function ZZCrdTyLvs$()
ZZCrdTyLvs = "1 2 3"
End Function

Private Sub CrdExpr__Tst()
Dim S$: S = CrdExpr(ZZCrdTyLvs)
Debug.Print RplVBar(S)
End Sub

Private Sub GpItm__Tst()
Debug.Print GpItm(5)
End Sub
