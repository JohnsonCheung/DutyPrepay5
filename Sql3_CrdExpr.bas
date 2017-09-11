Attribute VB_Name = "Sql3_CrdExpr"
Option Compare Database
Option Explicit
Const Pfx$ = "|    "

Sub AAAA()
Sql3CrdExpr__Tst
End Sub

Function Sql3CrdExpr$(CrdPfxTyDry() As Variant, CrdTyLvs$)
        Const CaseWhen$ = "    Case When"
    Const ElseCaseWhen$ = "|    Else Case When"
Dim CrdTyAy%(): CrdTyAy = AyAsgInto(SplitLvs(CrdTyLvs), EmptyIntAy)
Dim NGp%:           NGp = Sz(CrdTyAy)
Dim ElseN$:       ElseN = Pfx & "Else " & NGp + 1
Dim EndN$:         EndN = Pfx & StrDup(NGp, "End ")
Dim GpAy():        GpAy = AyMap(CrdTyAy, "GpItm", CrdPfxTyDry)
Dim Gp$:             Gp = Join(GpAy, ElseCaseWhen)
Sql3CrdExpr = CaseWhen & Gp & ElseN & EndN
End Function

Private Function CrdPfxAy(CrdTyId%, CrdPfxTyDry()) As String()
Dim Dry(): Dry = DrySel(CrdPfxTyDry, 1, CrdTyId)
CrdPfxAy = DryColStr(Dry, 0)
End Function

Private Function GpItm$(CrdTyId%, CrdPfxTyDry())
Dim Ay$(): Ay = SHMCodeLikAy(CrdTyId, CrdPfxTyDry)
Const Sep$ = " OR"
GpItm = Join(Ay, Sep) & " THEN " & CrdTyId
End Function

Private Function SHMCodeLik(Pfx)
SHMCodeLik = FmtQQ("|    SHMCode Like '?%'", Pfx)
End Function

Private Function SHMCodeLikAy(CrdTyId%, CrdPfxTyDry()) As String()
Dim O$(), Pfx
Dim PfxAy$(): PfxAy = CrdPfxAy(CrdTyId, CrdPfxTyDry)
For Each Pfx In PfxAy
    Push O, SHMCodeLik(Pfx)
Next
O = AyAlignL(O)
SHMCodeLikAy = O
End Function

Private Function ZZCrdPfxTyDry() As Variant()
Dim O()
Push O, Array("134234", 1)
Push O, Array("12323", 1)
Push O, Array("2444", 2)
Push O, Array("2443434", 2)
Push O, Array("24424", 2)
Push O, Array("3", 3)
Push O, Array("5446561", 4)
Push O, Array("6234341", 5)
Push O, Array("6234342", 5)
ZZCrdPfxTyDry = O
End Function

Private Function ZZCrdTyLvs$()
ZZCrdTyLvs = "1 2 3"
End Function

Private Sub GpItm__Tst()
Debug.Print GpItm(5, ZZCrdPfxTyDry)
End Sub

Private Sub Sql3CrdExpr__Tst()
Dim S$: S = Sql3CrdExpr(ZZCrdPfxTyDry, ZZCrdTyLvs)
Debug.Print RplVBar(S)
End Sub
