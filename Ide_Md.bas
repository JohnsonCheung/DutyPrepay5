Attribute VB_Name = "Ide_Md"
Option Explicit
Option Compare Database

Function DftMd(Optional A As CodeModule) As CodeModule
If IsNothing(A) Then
    Set DftMd = Application.Vbe.ActiveCodePane.CodeModule
Else
    Set DftMd = A
End If
End Function

Function DftMdNm$(Nm)
Dim N$:
    If Not IsMissing(Nm) Then N = Nm
If N = "" Then
    DftMdNm = MdNm(DftMd)
Else
    DftMdNm = N
End If
End Function

Function Md(Optional MdNm, Optional A As Vbproject) As CodeModule
Set Md = DftPj(A).VBComponents(DftMdNm(MdNm)).CodeModule
End Function

Sub MdAppDclLin(DclLin$, Optional A As CodeModule)
With DftMd(A)
    .InsertLines Md.CountOfDeclarationLines + 1, DclLin
End With
Debug.Print FmtQQ("MdAppDclLin: Module(?) a DclLin is inserted", MdNm(A))
End Sub

Sub MdAppLines(Lines$, Optional A As CodeModule)
If Lines = "" Then Exit Sub
Dim M As CodeModule
    Set M = DftMd(A)
Dim Bef%
    Bef = M.CountOfLines
M.InsertLines MdLasLno(A) + 1, Lines
Dim Aft%
    Aft = M.CountOfLines
Dim Exp%
    Exp = Bef + LinesLinCnt(Lines)
If Exp <> Aft Then Debug.Print FmtQQ("MdAppLines Er(LinCnt Added is not expected): Bef[?] LinCnt[?]: Exp(Bef+LinCnt)[?] <> Aft[?] AftBdyLinCnt[?]", Bef, LinesLinCnt(Lines), Exp, Aft, LinesLinCnt(MdBdyLines(A)))
End Sub

Sub MdAppLy(Ly$(), Optional A As CodeModule)
MdAppLines JnCrLf(Ly), A
End Sub

Function MdBdyLines$(Optional A As CodeModule)
MdBdyLines = SrcBdyLines(MdSrc(A))
End Function

Function MdBdyLnoCnt(Optional A As CodeModule) As LnoCnt
MdBdyLnoCnt = SrcBdyLnoCnt(MdSrc(A))
End Function

Function MdBdyLy(Optional A As CodeModule) As String()
MdBdyLy = SrcBdyLy(MdSrc(A))
End Function

Function MdCanHasCd(Optional A As CodeModule) As Boolean
Select Case MdTy(A)
Case _
    vbext_ComponentType.vbext_ct_StdModule, _
    vbext_ComponentType.vbext_ct_ClassModule, _
    vbext_ComponentType.vbext_ct_Document, _
    vbext_ComponentType.vbext_ct_MSForm
    MdCanHasCd = True
End Select
End Function

Sub MdClr(Optional A As CodeModule)
With DftMd(A)
    If .CountOfLines = 0 Then Exit Sub
    Debug.Print FmtQQ("MdClr: Md(?) of lines(?) is cleared", MdNm(A), .CountOfLines)
    .DeleteLines 1, .CountOfLines
End With
End Sub

Function MdCmp(Optional A As CodeModule) As VBComponent
Set MdCmp = DftMd(A).Parent
End Function

Function MdCmpTy(Optional A As CodeModule) As vbext_ComponentType
MdCmpTy = MdCmp(A).Type
End Function

Function MdContLin$(Lno&, Optional A As CodeModule)
Dim J&, L&, Md As CodeModule: Set Md = A
L = Lno
Dim O$: O = Md.Lines(L, 1)
While LasChr(O) = "_"
    L = L + 1
    O = RmvLasChr(O) & Md.Lines(L, 1)
Wend
MdContLin = O
End Function

Sub MdCpy(Fm$, ToMdNm$, Optional A As Vbproject)
Dim FmMd As CodeModule: Set FmMd = Md(Fm, A)
Dim Ty As vbext_ComponentType: Ty = MdTy(FmMd)
MdNew ToMdNm, Ty, A
Dim O As CodeModule: Set O = Md(ToMdNm)
MdAppLy MdLy(FmMd), O
End Sub

Function MdDclLines$(Optional A As CodeModule)
MdDclLines = JnCrLf(MdDclLy(A))
End Function

Function MdDclLy(Optional A As CodeModule) As String()
MdDclLy = SrcDclLy(MdSrc(A))
End Function

Function MdEnmCnt%(Optional A As CodeModule)
MdEnmCnt = SrcEnmCnt(MdDclLy(A))
End Function

Function MdEns(MdNm, Optional Ty As vbext_ComponentType = vbext_ct_StdModule, Optional A As Vbproject) As CodeModule
If Not PjHasMd(MdNm, A) Then
    MdNew MdNm, Ty, A
End If
Set MdEns = Md(MdNm, A)
End Function

Function MdEnsMth(MthNm$, NewMthLines$, Optional A As CodeModule)
Dim OldMthLines$: OldMthLines = MdMthLines(MthNm, A)
If OldMthLines = NewMthLines Then
    Debug.Print FmtQQ("MdEnsMth: Mth(?) in Md(?) is same", MthNm, MdNm(A))
End If
MdRmvMth MthNm, A
MdAppLines NewMthLines, A
Debug.Print FmtQQ("MdEnsMth: Mth(?) in Md(?) is replaced <=========", MthNm, MdNm(A))
End Function


Function MdExp(Optional A As CodeModule)
Dim Md As CodeModule: Set Md = DftMd(A)
Md.Parent.Export MdSrcFfn(Md)
Debug.Print MdNm(A)
End Function

Sub MdGo(MdNm$, Optional ClsOth As Boolean, Optional A As Vbproject)
If ClsOth Then WinClsCd
Md(MdNm, A).CodePane.Show
End Sub

Function MdInfDr(Optional A As CodeModule) As Variant()
Dim Pj$, Md$, IsCls As Boolean, LinCnt%, FunCnt%, TyCnt%, EnmCnt%
Dim M As CodeModule: Set M = DftMd(A)
Pj = PjNm(MdPj(M))
Md = MdNm(M)
LinCnt = M.CountOfLines
FunCnt = MdMthCnt(M)
TyCnt = MdTyCnt(M)
EnmCnt = MdEnmCnt(M)
IsCls = MdIsCls(M)
MdInfDr = Array(Pj, Md, IsCls, LinCnt, FunCnt, TyCnt, EnmCnt)
End Function

Function MdInfDry(Optional A As Vbproject) As Variant()
Dim I, M As CodeModule, O()
For Each I In PjMdAy(, A)
    Set M = I
    Push O, MdInfDr(M)
Next
MdInfDry = O
End Function

Function MdInfDt(Optional A As Vbproject) As Dt
Dim O As Dt
O.DtNm = "MdInf"
O.Fny = SplitSpc("Pj Md IsCls LinCnt FunCnt TyCnt EnmCnt")
O.Dry = MdInfDry(A)
MdInfDt = O
End Function

Sub MdInsLines(Lno&, Lines$, Optional A As CodeModule)
With DftMd(A)
    .InsertLines Lno, Lines
End With
End Sub

Function MdIsCls(Optional A As CodeModule) As Boolean
MdIsCls = MdTy(A) = vbext_ct_ClassModule
End Function

Function MdIsEmpty(Optional A As CodeModule) As Boolean
MdIsEmpty = DftMd(A).CountOfLines = 0
End Function

Function MdIsExist(MdNm$, Optional A As Vbproject) As Boolean
On Error GoTo X
MdIsExist = DftPj(A).VBComponents(MdNm).Name = MdNm
Exit Function
X:
End Function

Function MdLasLno&(Optional A As CodeModule)
MdLasLno = DftMd(A).CountOfLines
End Function

Function MdLin$(Lno&, Optional A As CodeModule)
If Lno <= 0 Then Exit Function
With DftMd(A)
    If Lno <= .CountOfLines Then MdLin = .Lines(Lno, 1)
End With
End Function

Function MdLines$(Optional A As CodeModule)
With DftMd(A)
    If .CountOfLines = 0 Then Exit Function
    MdLines = .Lines(1, .CountOfLines)
End With
End Function

Function MdLinesByLnoCnt$(LnoCnt As LnoCnt, Optional A As CodeModule)
With LnoCnt
    If .Cnt <= 0 Then Exit Function
    MdLinesByLnoCnt = DftMd(A).Lines(.Lno, .Cnt)
End With
End Function

Function MdLnoCntLines$(LnoCnt As LnoCnt, Optional A As CodeModule)
With LnoCnt
    If .Cnt = 0 Then Exit Function
    MdLnoCntLines = DftMd(A).Lines(.Lno, .Cnt)
End With
End Function

Function MdLy(Optional A As CodeModule) As String()
MdLy = SplitCrLf(MdLines(A))
End Function

Function MdMthCnt%(Optional A As CodeModule)
MdMthCnt = SrcMthCnt(MdSrc(A))
End Function

Function MdMthDrs(Optional WithBdyLy As Boolean, _
    Optional WithBdyLines As Boolean, Optional A As CodeModule) As Drs
Dim O As Drs
    O = SrcMthDrs(MdSrc(A), WithBdyLy, WithBdyLines)
MdMthDrs = DrsAddConstCol(O, "MdNm", MdNm(A))
End Function

Function MdMthLin$(Optional MthNm$, Optional A As CodeModule)
MdMthLin = SrcMthLin(MdBdyLy(A), DftMthNm(MthNm))
End Function

Function MdMthLines$(MthNm$, Optional A As CodeModule)
MdMthLines = SrcMthLines(MdSrc(A), MthNm)
End Function

Function MdMthLno&(Optional MthNm, Optional A As CodeModule)
MdMthLno = 1 + SrcMthIdx(MdSrc(A), DftMthNm(MthNm))
End Function

Function MdMthLnoAy(MthNm, Optional A As CodeModule)
MdMthLnoAy = AyIncN(SrcMthIdxAy(MdSrc(A), MthNm))
End Function

Function MdMthLnoCntAy(MthNm$, Optional A As CodeModule) As LnoCnt()
MdMthLnoCntAy = SrcMthFmToAy(MdSrc(A), MthNm)
End Function

Function MdMthLy(Optional MthNm$, Optional A As CodeModule) As String()
MdMthLy = SrcMthLy(MdSrc(A), DftMthNm(MthNm))
End Function

Function MdMthNy(Optional MthNmLik = "*", Optional A As CodeModule) As String()
MdMthNy = SrcMthNy(MdSrc(A), MthNmLik)
End Function

Sub MdNew(Optional MdNm, Optional Ty As vbext_ComponentType = vbext_ct_StdModule, Optional A As Vbproject)
Dim O As VBComponent: Set O = DftPj(A).VBComponents.Add(Ty)
If MdNm <> "" Then O.Name = MdNm
MdEnsOptExplicit O.CodeModule
End Sub

Function MdNm(Optional A As CodeModule)
MdNm = DftMd(A).Parent.Name
End Function

Function MdPj(Optional A As CodeModule) As Vbproject
Set MdPj = DftMd(A).Parent.Collection.Parent
End Function

Function MdPrvMthNy(Optional A As CodeModule) As String()
MdPrvMthNy = SrcPrvMthNy(MdSrc(A))
End Function

Sub MdRmv(MdNm$)
Dim C As VBComponent: Set C = Md(MdNm).Parent
C.Collection.Remove C
End Sub

Sub MdRmvBdy(A As CodeModule)
MdRmvLnoCnt MdBdyLnoCnt(A), A
End Sub

Sub MdRmvLnoCnt(LnoCnt As LnoCnt, A As CodeModule)
With LnoCnt
    If .Cnt = 0 Then Exit Sub
    DftMd(A).DeleteLines .Lno, .Cnt
End With
End Sub

Sub MdRmvLnoCntAy(LnoCntAy() As LnoCnt, A As CodeModule)
Dim J%
For J = 0 To LnoCntUB(LnoCntAy)
    MdRmvLnoCnt LnoCntAy(J), A
Next
End Sub

Sub MdRmvMth(MthNm$, Optional A As CodeModule)
Dim M() As LnoCnt: M = MdMthLnoCntAy(MthNm, A)
If LnoCntAyIsEmpty(M) Then
    Debug.Print FmtQQ("Fun[?] in Md[?] not found, cannot Rmv", MthNm, MdNm(A))
Else
    Debug.Print FmtQQ("Fun[?] in Md[?] is removed", MthNm, MdNm(A))
End If
MdRmvLnoCntAy M, A
End Sub

Sub MdRmvTstMth(Optional A As CodeModule)
MdRmvMth "Tst", A
End Sub

Sub MdRpl(NewMdLines$, Optional A As CodeModule)
MdClr A
DftMd(A).InsertLines 1, NewMdLines
End Sub

Sub MdRplLin(Lno&, NewLin$, Optional A As CodeModule)
With DftMd(A)
    .DeleteLines Lno
    .InsertLines Lno, NewLin
End With
End Sub

Function MdSrc(Optional A As CodeModule) As String()
MdSrc = MdLy(A)
End Function

Function MdSrcFfn$(Optional A As CodeModule)
MdSrcFfn = PjSrcPth(MdPj(A)) & MdSrcFn(A)
End Function

Function MdSrcFn$(Optional A As CodeModule)
MdSrcFn = MdCmp(A).Name & MdSrcExt(A)
End Function

Sub MdSrt(Optional A As CodeModule)
Dim Md As CodeModule: Set Md = DftMd(A)
If MdNm(Md) = "Ide" Then
    Debug.Print "Ide", "<<<< Skipped"
    Exit Sub
End If
Dim Old$: Old = MdBdyLines(Md)
Dim NewLines$: NewLines = MdSrtedBdyLines(Md)
If Old = NewLines Then
    Debug.Print MdNm(Md), "<== Same"
    Exit Sub
End If
Debug.Print MdNm(Md), "<-- Sorted"
MdRmvBdy Md
MdAppLines NewLines, Md
End Sub

Function MdSrtedBdyLines$(Optional A As CodeModule)
If MdIsEmpty(A) Then Exit Function
Dim Drs As Drs
    Drs = SrcMthDrs(MdSrc(A), WithBdyLines:=True)
Dim MthLinesAy$()
    MthLinesAy = DrsStrCol(Drs, "BdyLines")
Dim I&()
    Dim Ky$(): Ky = MthDrsKy(Drs)
    I = AySrtIntoIdxAy(Ky)
Dim O$()
Dim J%
    For J = 0 To UB(I)
        Push O, vbCrLf & MthLinesAy(I(J))
    Next
MdSrtedBdyLines = JnCrLf(O)
End Function

Sub MdSrtRpt(Optional A As CodeModule)
Dim Md As CodeModule: Set Md = DftMd(A)
Dim Old$: Old = MdBdyLines(Md)
Dim NewLines$: NewLines = MdSrtedBdyLines(Md)
Dim O$: O = IIf(Old = NewLines, "(Same)", "<====Diff")
Debug.Print MdNm(Md), O
End Sub

Sub MdSrtRptDif(Optional A As CodeModule)
Dim Md As CodeModule: Set Md = DftMd(A)
Dim Old$: Old = MdBdyLines(Md)
Dim NewLines$: NewLines = MdSrtedBdyLines(Md)
If Old <> NewLines Then
    Debug.Print MdNm(Md), "<==== Dif"
End If
End Sub

Function MdTstMthLines(Optional A As CodeModule)
MdTstMthLines = JnCrLf(MdTstMthLy(A))
End Function

Function MdTstMthLy(Optional A As CodeModule) As String()
Dim O$(), Ay$()
Ay = MdTstMthNy(A)
If AyIsEmpty(Ay) Then Exit Function
Push O, "Sub Tst()"
PushAy O, AySrt(MdTstMthNy(A))
Push O, "End Sub"
MdTstMthLy = O
End Function

Function MdTstMthNy(Optional A As CodeModule) As String()
If MdIsEmpty(A) Then Exit Function
Dim M As Drs: M = MdMthDrs(A:=A)
Dim Dr
Dim O$(), Mdy$, Ty$, MthNm$, IMthNm%
Fiy M.Fny, "MthNm", IMthNm
If AyIsEmpty(M.Dry) Then Exit Function
For Each Dr In M.Dry
    MthNm = Dr(IMthNm)
    If IsSfx(MthNm, "__Tst") Then
        Push O, MthNm
    End If
Next
MdTstMthNy = O
End Function

Function MdTstMthNy_WithEr(Optional A As CodeModule) As String()
Dim M As Drs: M = MdMthDrs(A:=A)
Dim O$(), Mdy$, Ty$, MthNm$, Dr
Dim IMdy%, ITy%, IMthNm%
Fiy M.Fny, "Mdy Ty MthNm", IMdy, ITy, IMthNm
If AyIsEmpty(M.Dry) Then Exit Function
For Each Dr In M.Dry
    If IsSfx(Dr(IMthNm), "__Tst") Then
        If Dr(IMdy) <> "Private" Or Dr(ITy) <> "Sub" Then
            Push O, Dr(IMthNm)
        End If
    End If
Next
MdTstMthNy_WithEr = O
End Function

Function MdTy(Optional A As CodeModule) As vbext_ComponentType
MdTy = DftMd(A).Parent.Type
End Function

Function MdTyCnt%(Optional A As CodeModule)
MdTyCnt = SrcTyCnt(MdDclLy(A))
End Function

Sub MdUpdTstMth(Optional A As CodeModule)
If Not MdCanHasCd(A) Then Exit Sub
    
Dim NewLines$: NewLines = MdTstMthLines(A)
Dim OldLines$: OldLines = MdMthLines("Tst", A)
If OldLines = NewLines Then
    Debug.Print FmtQQ("Fun[Tst] in Md[?] is same: [?] lines", MdNm(A), LinesLinCnt(OldLines))
    Exit Sub
End If

MdRmvTstMth A
Dim Ly$()
    Ly = MdTstMthLy(A)

If Sz(Ly) > 0 Then
    Debug.Print FmtQQ("Fun[Tst] in Md[?] is inserted", MdNm(A))
    MdAppLy Ly, A
End If
End Sub

Private Function MdSrcExt$(Optional A As CodeModule)
Dim O$
Select Case MdCmpTy(A)
Case vbext_ct_ClassModule: O = ".cls"
Case vbext_ct_Document: O = ".cls"
Case vbext_ct_StdModule: O = ".bas"
Case Else: Err.Raise 1, , "MdSrcExt: Unexpected MdCmpTy.  Should be [Class or Module or Document]"
End Select
MdSrcExt = O
End Function

Private Sub MdAppLines__Tst()
Const MdNm$ = "Module1"
Dim M As CodeModule
    Set M = Md(MdNm)
MdAppLines "'aa", M
End Sub

Function MdInfDt__Tst()
DtBrw MdInfDt
End Function

Private Sub MdLy__Tst()
AyBrw MdLy
End Sub

Private Sub MdMthDrs__Tst()
DrsBrw MdMthDrs(WithBdyLy:=True, A:=Md("Vb_Str"))
End Sub

Private Sub MdMthLines__Tst()
Debug.Print Len(MdMthLines("MdMthLines"))
Debug.Print MdMthLines("MdMthLines")
End Sub

Private Sub MdSrt__Tst()
MdSrt Md("bb_Lib_Acs")
End Sub

Private Sub MdSrtedBdyLines__Tst()
StrBrw MdSrtedBdyLines(Md("Vb_Str"))
End Sub

Private Sub MdTstMthLines__Tst()
Dim Ny$()
    Ny = PjMdNy
    Ny = AySrt(Ny)
    
Dim I, M As CodeModule

Dim Dr(), Dry()
For Each I In Ny
    Set M = Md(I)
    Dr = Array(I, MdTstMthLines(M))
    Push Dry, Dr
Next
Dim Drs As Drs
    Drs.Dry = Dry
    Drs.Fny = SplitSpc("Md TstMth")
    Drs = DrsExpLinesCol(Drs, "TstMth")
DrsBrw Drs, , "Md"
End Sub

