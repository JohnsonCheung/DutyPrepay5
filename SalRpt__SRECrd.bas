Attribute VB_Name = "SalRpt__SRECrd"
Option Compare Database
Option Explicit

Function SRECrd$(CrdTyLvs$, CrdPfxTyDry())
Const CSub$ = "SECrd"
'SRECrd = Sales-Report-Expression-Card
Const CaseWhen$ = "Case When"
Const ElseCaseWhen$ = "|Else Case When"
Dim CrdTyAy%()
    CrdTyAy = AyIntAy(SplitLvs(CrdTyLvs))
    Dim StdCrdTyAy%()
    StdCrdTyAy = DrySelDisIntCol(CrdPfxTyDry, 1) ' 1 is colidx which is CrdTyId
    Dim NotExistIdAy%()
        NotExistIdAy = AyMinus(CrdTyAy, StdCrdTyAy)
        If Not AyIsEmpty(NotExistIdAy) Then
            Er CSub, "{CrdTyLvs} has item not found in std {CrdTyAy} which is comming from {CrdPfxTyDry}"
        End If
    If AyIsEmpty(CrdTyAy) Then CrdTyAy = StdCrdTyAy
Dim NGp%, GpAy$()
    NGp = Sz(CrdTyAy)
    Dim O$(), J%
    For J = 0 To UB(CrdTyAy)
        Push O, GpItm(CrdTyAy(J), CrdPfxTyDry)
    Next
    GpAy = O

Dim Gp$, ElseN$, EndN$
    ElseN = "|Else " & NGp + 1
    EndN = "|" & StrDup(NGp, "End ")
    Gp = Join(GpAy, ElseCaseWhen)
SRECrd = CaseWhen & Gp & ElseN & EndN
End Function

Private Function GpItm$(CrdTyId%, CrdPfxTyDry())
Dim Ay$(): Ay = SHMCodeLikAy(CrdTyId, CrdPfxTyDry)
Const Sep$ = " OR"
GpItm = Join(Ay, Sep) & " THEN " & CrdTyId
End Function

Private Function SHMCodeLikAy(CrdTyId%, CrdPfxTyDry()) As String()
Dim CrdPfxAy() As String
    Dim Dry(): Dry = DrySel(CrdPfxTyDry, 1, CrdTyId)
    CrdPfxAy = DryStrCol(Dry, 0)

Dim O$(), Pfx
    Dim SHMCodeLik$
    For Each Pfx In CrdPfxAy
        SHMCodeLik = FmtQQ("|SHMCode Like '?%'", Pfx)
        Push O, SHMCodeLik
    Next
O = AyAlignL(O)
SHMCodeLikAy = O
End Function
