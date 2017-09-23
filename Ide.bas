Attribute VB_Name = "Ide"
Option Explicit
Option Compare Database

Private Enum AA 'For testing
    AA1
    '
    
End Enum

Type SrcLinBrk
    MthNm As String
    Ty As String    ' Sub Function Get Set Let (Ty here means SrcTy)
    Mdy As String
End Type
Type LnoCnt
    Lno As Long
    Cnt As Long
End Type

Type FmTo
    FmIdx As Long
    ToIdx As Long
End Type
Sub LnoCntPush(O() As LnoCnt, M As LnoCnt)
Dim N&: N = LnoCntSz(O)
ReDim Preserve O(N)
O(N) = M
End Sub
Function LnoCntAyIsEmpty(A() As LnoCnt) As Boolean
LnoCntAyIsEmpty = LnoCntSz(A) = 0
End Function
Function LnoCntUB&(A() As LnoCnt)
LnoCntUB = LnoCntSz(A) - 1
End Function
Function LnoCntSz&(A() As LnoCnt)
On Error Resume Next
LnoCntSz = UBound(A) + 1
End Function
Function DftMd(Optional A As CodeModule) As CodeModule
If IsNothing(A) Then
    Set DftMd = Application.VBE.ActiveCodePane.CodeModule
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

Function DftPj(Optional A As VBProject) As VBProject
If IsNothing(A) Then
    Set DftPj = Application.VBE.ActiveVBProject
Else
    Set DftPj = A
End If
End Function

Function EnmBdyLy(EnmNm$, Optional A As CodeModule) As String()
Dim B%: B = EnmLinIdx(EnmNm, A): If B = -1 Then Exit Function
Dim O$(), Ly$(), J%
Ly = MdDclLy(A)
For J = B To UB(Ly)
    Push O, Ly(J)
    If IsPfx(Ly(J), "End Enum") Then EnmBdyLy = O: Exit Function
Next
Stop
End Function

Function EnmIsMbrLin(L) As Boolean
If SrcLinIsRmk(L) Then Exit Function
If Trim(L) = "" Then Exit Function
EnmIsMbrLin = True
End Function

Function EnmLinIdx%(EnmNm$, Optional A As CodeModule)
Dim Ly$(): Ly = MdDclLy(A)
Dim U%: U = UB(Ly)
Dim O%, L$
For O = 0 To U
    If SrcLinIsEnm(Ly(O)) Then
        L = Ly(O)
        ParseMdy L
        L = RmvFstTerm(L)
        If FstTerm(L) = EnmNm Then
            EnmLinIdx = O: Exit Function
        End If
    End If
Next
EnmLinIdx = -1
End Function

Function EnmMbrCnt%(EnmNm$, Optional A As CodeModule)
EnmMbrCnt = Sz(EnmMbrLy(EnmNm, A))
End Function

Function EnmMbrLy(EnmNm$, Optional A As CodeModule) As String()
Dim Ly$(), O$(), J%
Ly = EnmBdyLy(EnmNm, A)
For J = 1 To UB(Ly) - 1
    If EnmIsMbrLin(Ly(J)) Then Push O, Ly(J)
Next
EnmMbrLy = O
End Function
Function FmToSz&(A() As FmTo)
On Error Resume Next
FmToSz = UBound(A) + 1
End Function
Function FmToUB&(A() As FmTo)
FmToUB = FmToSz(A) - 1
End Function
Sub FmToPush(O() As FmTo, M As FmTo)
Dim N&: N = FmToSz(O)
ReDim Preserve O(N)
O(N) = M
End Sub
Function FmToAyIsEmpty(A() As FmTo) As Boolean
FmToAyIsEmpty = FmToSz(A) = 0
End Function
Function FmToAyLnoCntAy(A() As FmTo) As LnoCnt()
If FmToAyIsEmpty(A) Then Exit Function
Dim U&, J&
U = FmToUB(A)
Dim O() As LnoCnt
    ReDim O(U)
For J = 0 To U
    O(J) = FmToLnoCnt(A(J))
Next
FmToAyLnoCntAy = O
End Function

Function FmTo(FmIdx&, ToIdx&) As FmTo
FmTo.FmIdx = FmIdx
FmTo.ToIdx = ToIdx
End Function

Function FmToLnoCnt(A As FmTo) As LnoCnt
Dim Lno&, Cnt&
    Cnt = A.ToIdx - A.FmIdx + 1
    If Cnt < 0 Then Cnt = 0
    Lno = A.FmIdx + 1
With FmToLnoCnt
    .Cnt = Cnt
    .Lno = Lno
End With
End Function

Function MthDrsKy(MthDrs As Drs) As String()
Dim Dry() As Variant: Dry = MthDrs.Dry
Dim Fny$(): Fny = MthDrs.Fny
Dim O$()
    Dim Ty$, Mdy$, MthNm$, K$, IdxAy&(), Dr
    IdxAy = FnyLIdxAy(Fny, "Mdy MthNm Ty")
    If AyIsEmpty(MthDrs.Dry) Then Exit Function
    For Each Dr In MthDrs.Dry
        'Debug.Print Mdy, MthNm, Ty
        AyAsg_Idx Dr, IdxAy, Mdy, MthNm, Ty
        Push O, MthKey(Mdy, Ty, MthNm)
    Next
MthDrsKy = O
End Function

Function MthKey$(Mdy$, Ty$, MthNm$)
Dim A1 As Byte
    If IsSfx(MthNm, "__Tst") Then
        A1 = 8
    ElseIf MthNm = "Tst" Then
        A1 = 9
    Else
        Select Case Mdy
        Case "Public", "": A1 = 1
        Case "Friend": A1 = 2
        Case "Private": A1 = 3
        Case Else: Stop
        End Select
    End If
Dim A3$
    If Ty <> "Function" And Ty <> "Sub" Then A3 = Ty
MthKey = FmtQQ("?:?:?", A1, MthNm, A3)
End Function

Function SrcLinEndLinPfx$(SrcLin)
Dim A$: A = SrcLinMthTy(SrcLin): If A = "" Then Stop
SrcLinEndLinPfx = "End " & A
End Function

Function IsEmptyMd(Optional A As CodeModule) As Boolean
IsEmptyMd = DftMd(A).CountOfLines = 0
End Function

Function JnContinueLin(Ly$()) As String()
Dim O$(): O = Ly
Dim J&
For J = UB(O) - 1 To 0 Step -1
    If LasChr(O(J)) = "_" Then
        O(J) = RmvLasNChr(O(J)) & O(J + 1)
        O(J + 1) = ""
    End If
Next
JnContinueLin = O
End Function

Function SrcLinIsCd(Lin) As Boolean
Dim L$: L = Trim(Lin)
If L = "" Then Exit Function
If FstChr(L) = "'" Then Exit Function
SrcLinIsCd = True
End Function

Function LnoCnt(Lno&, Cnt&) As LnoCnt
LnoCnt.Lno = Lno
LnoCnt.Cnt = Cnt
End Function

Function LnoCntStr$(A As LnoCnt)
LnoCntStr = FmtQQ("Lno(?) Cnt(?)", A.Lno, A.Cnt)
End Function

Function Md(Optional MdNm, Optional A As VBProject) As CodeModule
Set Md = DftPj(A).VBComponents(DftMdNm(MdNm)).CodeModule
End Function

Function MdAddOptExplicit(Optional A As CodeModule)
DftMd(A).InsertLines 1, "Option Explicit"
Debug.Print MdNm(A), "<-- Option Explicit added"
End Function

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

Sub MdCpy(Fm$, ToMdNm$, Optional A As VBProject)
Dim FmMd As CodeModule: Set FmMd = Md(Fm, A)
Dim Ty As vbext_ComponentType: Ty = MdTy(FmMd)
MdNew ToMdNm, Ty, A
Dim O As CodeModule: Set O = Md(ToMdNm)
MdAppLy MdLy(FmMd), O
End Sub

Function MdDclLines$(Optional A As CodeModule)
With DftMd(A)
    If .CountOfDeclarationLines = 0 Then Exit Function
    MdDclLines = .Lines(1, .CountOfDeclarationLines)
End With
End Function

Function MdDclLy(Optional A As CodeModule) As String()
MdDclLy = SplitCrLf(MdDclLines(A))
End Function

Function MdEnmCnt%(Optional A As CodeModule)
MdEnmCnt = SrcEnmCnt(MdDclLy(A))
End Function

Sub MdEnsOptExplicit(Optional A As CodeModule)
If Not MdHasOptExplicit(A) Then MdAddOptExplicit A
'If MdHasOptExplicit(A) Then
'    Debug.Print MdNm(A), "(* With Option Explicit *)"
'Else
'    Debug.Print MdNm(A), "<-------------------- No Option Explicit"
'End If
End Sub

Function MdExp(Optional A As CodeModule)
Dim Md As CodeModule: Set Md = DftMd(A)
Md.Parent.Export MdSrcFfn(Md)
Debug.Print MdNm(A)
End Function

Function MdMthCnt%(Optional A As CodeModule)
MdMthCnt = SrcMthCnt(MdSrc(A))
End Function

Function MdMthDrs(Optional WithBdyLy As Boolean, Optional WithBdyLines As Boolean, Optional A As CodeModule) As Drs
MdMthDrs = SrcMthDrs(MdSrc(A), MdNm(A), WithBdyLy, WithBdyLines)
End Function

Function MdMthLines$(MthNm$, Optional A As CodeModule)
MdMthLines = SrcMthLines(MdSrc(A), MthNm)
End Function

Function MdMthLnoCntAy(MthNm$, Optional A As CodeModule) As LnoCnt()
MdMthLnoCntAy = SrcMthFmToAy(MdSrc(A), MthNm)
End Function

Function MdMthNy(Optional A As CodeModule) As String()
MdMthNy = SrcMthNy(MdSrc(A))
End Function

Function MdHasOptExplicit(Optional A As CodeModule)
Dim Ay$()
    Ay = MdDclLy(A)
Dim I
For Each I In Ay
    If I = "Option Explicit" Then MdHasOptExplicit = True: Exit Function
Next
End Function

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

Function MdInfDry(Optional A As VBProject) As Variant()
Dim I, M As CodeModule, O()
For Each I In PjMdAy(A)
    Set M = I
    Push O, MdInfDr(M)
Next
MdInfDry = O
End Function

Function MdInfDt(Optional A As VBProject) As Dt
Dim O As Dt
O.DtNm = "MdInf"
O.Fny = SplitSpc("Pj Md IsCls LinCnt FunCnt TyCnt EnmCnt")
O.Dry = MdInfDry(A)
MdInfDt = O
End Function

Function MdIsCls(Optional A As CodeModule) As Boolean
MdIsCls = MdTy(A) = vbext_ct_ClassModule
End Function

Function MdIsEmpty(Optional A As CodeModule)
MdIsEmpty = (DftMd(A).CountOfLines = 0)
End Function

Function MdIsExist(MdNm$, Optional A As VBProject) As Boolean
On Error GoTo X
MdIsExist = DftPj(A).VBComponents(MdNm).Name = MdNm
Exit Function
X:
End Function

Function MdLasLno&(Optional A As CodeModule)
MdLasLno = DftMd(A).CountOfLines
End Function

Function MdLines$(Optional A As CodeModule)
With DftMd(A)
    If .CountOfLines = 0 Then Exit Function
    MdLines = .Lines(1, .CountOfLines)
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

Function MdLy_Jn(Optional A As CodeModule) As String()
MdLy_Jn = JnContinueLin(MdLy(A))
End Function

Sub MdNew(Optional MdNm$, Optional Ty As vbext_ComponentType = vbext_ct_StdModule, Optional A As VBProject)
Dim O As VBComponent: Set O = DftPj(A).VBComponents.Add(Ty)
O.CodeModule.DeleteLines 1, 2
If MdNm <> "" Then O.Name = MdNm
End Sub

Function MdNm(Optional A As CodeModule)
MdNm = DftMd(A).Parent.Name
End Function

Function MdPj(Optional A As CodeModule) As VBProject
Set MdPj = DftMd(A).Parent.Collection.Parent
End Function

Sub MdRmv(MdNm$)
Dim C As VBComponent: Set C = Md(MdNm).Parent
C.Collection.Remove C
End Sub

Sub MdRmvBdy(A As CodeModule)
MdRmvLnoCnt MdBdyLnoCnt(A), A
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
Sub MdRmvLnoCntAy(LnoCntAy() As LnoCnt, A As CodeModule)
Dim J%
For J = 0 To LnoCntUB(LnoCntAy)
    MdRmvLnoCnt LnoCntAy(J), A
Next
End Sub

Sub MdRmvLnoCnt(LnoCnt As LnoCnt, A As CodeModule)
With LnoCnt
    If .Cnt = 0 Then Exit Sub
    DftMd(A).DeleteLines .Lno, .Cnt
End With
End Sub

Sub MdRmvTstSub(Optional A As CodeModule)
MdRmvMth "Tst", A
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

Sub MdSrt(Optional A As CodeModule)
Dim Md As CodeModule: Set Md = DftMd(A)
If MdNm(Md) = "Ide" Then
    Debug.Print "Ide", "<<<< Skipped"
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
    Drs = MdMthDrs(WithBdyLines:=True, A:=A)
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

Function MdTstSubLines$(Optional A As CodeModule)
MdTstSubLines = JnCrLf(MdTstSubLy(A))
End Function

Function MdTstSubLy(Optional A As CodeModule) As String()
Dim O$(), Ay$()
Ay = MdTstSubNy(A)
If AyIsEmpty(Ay) Then Exit Function
Push O, "Sub Tst()"
PushAy O, AySrt(MdTstSubNy(A))
Push O, "End Sub"
MdTstSubLy = O
End Function

Function MdTstSubNy(Optional A As CodeModule) As String()
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
MdTstSubNy = O
End Function

Function MdTstSubNy_WithEr(Optional A As CodeModule) As String()
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
MdTstSubNy_WithEr = O
End Function

Function MdTy(Optional A As CodeModule) As vbext_ComponentType
MdTy = DftMd(A).Parent.Type
End Function

Function MdTyCnt%(Optional A As CodeModule)
MdTyCnt = SrcTyCnt(MdDclLy(A))
End Function

Sub MdUpdTstSub(Optional A As CodeModule)
If Not MdCanHasCd(A) Then Exit Sub
    
Dim NewLines$: NewLines = MdTstSubLines(A)
Dim OldLines$: OldLines = MdMthLines("Tst", A)
If OldLines = NewLines Then
    Debug.Print FmtQQ("Fun[Tst] in Md[?] is same: [?] lines", MdNm(A), LinesLinCnt(OldLines))
    Exit Sub
End If

MdRmvTstSub A
Dim Ly$()
    Ly = MdTstSubLy(A)

If Sz(Ly) > 0 Then
    Debug.Print FmtQQ("Fun[Tst] in Md[?] is inserted", MdNm(A))
    MdAppLy Ly, A
End If
End Sub

Function ParseMthTy$(OLin$)
ParseMthTy = RTrim(ParseOneOf(OLin, SyOfMthTy))
End Function

Function ParsePrpTy$(OLin$)
ParsePrpTy = RTrim(ParseOneOf(OLin, SyOfPrpTy))
End Function

Function ParseMdy$(OLin$)
ParseMdy = RTrim(ParseOneOf(OLin, SyOfMdy))
End Function

Function ParseNm$(OLin$)
Dim J%
J = 1
If Not IsLetter(FstChr(OLin)) Then GoTo Nxt
For J = 2 To Len(OLin)
    If Not IsNmChr(Mid(OLin, J, 1)) Then GoTo Nxt
Next
Nxt:
If J = 1 Then Exit Function
ParseNm = Left(OLin, J - 1)
OLin = Mid(OLin, J)
End Function

Function ParseOneOf(OLin$, OneOfAy$())
Dim I
For Each I In OneOfAy
    If IsPfx(OLin, I) Then OLin = RmvPfx(OLin, I): ParseOneOf = I: Exit Function
Next
End Function

Function Pj(PjNm) As VBProject
Set Pj = Application.VBE.VBProjects(PjNm)
End Function

Sub PjAssertNotUnderSrc(Optional A As VBProject)
Dim B$: B = PjPth(A)
If PthFdr(B) = "Src" Then Stop
End Sub

Sub PjCpyToSrc(Optional A As VBProject)
FilCpyToPth DftPj(A).FileName, PjSrcPth(A), OvrWrt:=True
End Sub

Sub PjEnsOptExplicit(Optional A As VBProject)
Dim I, Md As CodeModule
For Each I In PjMdAy(A)
    Set Md = I
    MdEnsOptExplicit Md
Next
End Sub

Sub PjExp(Optional A As VBProject)
PjAssertNotUnderSrc
PjCpyToSrc A
PthClrFil PjSrcPth(A)
Dim Md As CodeModule, I
For Each I In PjMdAy(A)
    Set Md = I
    MdExp Md
Next
End Sub

Function PjMthDrs(Optional WithBdyLy As Boolean, Optional WithBdyLines As Boolean, Optional A As VBProject) As Drs
Dim Dry()
    Dim I, Md As CodeModule
    For Each I In PjMdAy(A)
        Set Md = I
        PushAy Dry, MdMthDrs(WithBdyLy, WithBdyLines, A:=Md).Dry
    Next
Dim O As Drs
    O.Fny = PjMthDrsFny(WithBdyLy, WithBdyLines)
    O.Dry = Dry
PjMthDrs = O
End Function

Function PjMthDrsFny(WithBdyLy As Boolean, WithBdyLines As Boolean) As String()
Dim O$(): O = SplitSpc("Lno Mdy Ty MthNm MdNm")
If WithBdyLy Then Push O, "BdyLy"
If WithBdyLines Then Push O, "BdyLines"
PjMthDrsFny = O
End Function

Function PjMdAy(Optional A As VBProject) As CodeModule()
Dim O() As CodeModule
Dim Cmp As VBComponent
For Each Cmp In DftPj(A).VBComponents
    PushObj O, Cmp.CodeModule
Next
PjMdAy = O
End Function

Function PjMdNy(Optional A As VBProject) As String()
PjMdNy = OyPrp(PjMdAy(A), "Name", EmptySy)
End Function

Function PjNm$(Optional A As VBProject)
PjNm = DftPj(A).Name
End Function

Function PjPth$(Optional A As VBProject)
PjPth = FfnPth(DftPj(A).FileName)
End Function

Sub PjSrcBrw()
PthBrw PjSrcPth
End Sub

Function PjSrcPth$(Optional A As VBProject)
Dim Ffn$: Ffn = DftPj(A).FileName
Dim Fn$: Fn = FfnFn(Ffn)
Dim O$:
O = FfnPth(DftPj(A).FileName) & "Src\": PthEns O
O = O & Fn & "\":                       PthEns O
PjSrcPth = O
End Function
Sub PjSrtRptDif(Optional A As VBProject)
Dim Md As CodeModule, I
For Each I In PjMdAy(A)
    Set Md = I
    MdSrtRptDif Md
Next
End Sub

Sub PjSrtRpt(Optional A As VBProject)
Dim Md As CodeModule, I
For Each I In PjMdAy(A)
    Set Md = I
    MdSrtRpt Md
Next
End Sub

Sub PjSrt(Optional A As VBProject)
Dim Md As CodeModule, I
For Each I In PjMdAy(A)
    Set Md = I
    If MdNm(Md) <> "Ide" Then
        MdSrt Md
    End If
Next
End Sub

Function PjTstSubNy_WithEr(Optional A As VBProject) As String()
Dim O$(), I, Md As CodeModule
For Each I In PjMdAy(A)
    Set Md = I
    PushAy O, AyAddPfx(MdTstSubNy_WithEr(Md), MdNm(Md) & ".")
Next
PjTstSubNy_WithEr = O
End Function

Sub PjUpdTstSub(Optional A As VBProject)
Dim I, Md As CodeModule
For Each I In PjMdAy(A)
    Set Md = I
    MdUpdTstSub Md
Next
End Sub

Function SrcBdyLines$(Src$())
SrcBdyLines = JnCrLf(SrcBdyLy(Src))
End Function

Function SrcBdyLnoCnt(Src$()) As LnoCnt
Dim Lno&
Dim Cnt&
    Lno = SrcDclCnt(Src) + 1
    Cnt = Sz(Src) - Lno + 1
Dim O As LnoCnt
    O.Lno = Lno
    O.Cnt = Cnt
SrcBdyLnoCnt = O
End Function

Function SrcBdyLy(Src$()) As String()
SrcBdyLy = AyFm(Src, SrcDclCnt(Src))
End Function

Function SrcDclCnt&(Src$())
Dim I&: I = SrcFstMthIdx(Src)
If I = -1 Then SrcDclCnt = Sz(Src): Exit Function
SrcDclCnt = SrcMthIdxAyFstRmkIdx(Src, I)
End Function

Function SrcEnmCnt%(Src$())
If AyIsEmpty(Src) Then Exit Function
Dim I, O%
For Each I In Src
    If SrcLinIsEnm(I) Then O = O + 1
Next
SrcEnmCnt = O
End Function

Function SrcFstMthIdx&(Src$())
Dim J%
For J = 0 To UB(Src)
    If SrcLinIsMth(Src(J)) Then
        SrcFstMthIdx = J
        Exit Function
    End If
Next
SrcFstMthIdx = -1
End Function

Function SrcMthCnt%(Src$())
If AyIsEmpty(Src) Then Exit Function
Dim I, O%
For Each I In Src
    If SrcLinIsMth(I) Then O = O + 1
Next
SrcMthCnt = O
End Function

Function SrcMthDrs(Src$(), MdNm$, Optional WithBdyLy As Boolean, Optional WithBdyLines As Boolean) As Drs
SrcMthDrs.Dry = SrcMthDry(Src, MdNm$, WithBdyLy, WithBdyLines)
SrcMthDrs.Fny = PjMthDrsFny(WithBdyLy, WithBdyLines)
End Function
Function SrcLinDr(SrcLin, MdNm$, Lno&) As Variant()
With SrcLinBrk(SrcLin)
    SrcLinDr = Array(Lno, .Mdy, .Ty, .MthNm, MdNm)
End With
End Function
Function SrcMthDry(Src$(), MdNm$, Optional WithBdyLy As Boolean, Optional WithBdyLines As Boolean) As Variant()
Dim MthIdxAy&(): MthIdxAy = SrcMthIdxAyAll(Src)
Dim O()
    Dim Dr()
    Dim J&
    Dim MthIdx&
    Dim BdyLy$()
    For J = 0 To UB(MthIdxAy)
        MthIdx = MthIdxAy(J)
        Dr = SrcLinDr(Src(MthIdx), MdNm, MthIdx + 1)
        If WithBdyLy Or WithBdyLines Then
            BdyLy = SrcMthIdxBdyLy(Src, MthIdx)
            If WithBdyLy Then Push Dr, BdyLy
            If WithBdyLines Then Push Dr, JnCrLf(BdyLy)
        End If
        Push O, Dr
    Next
SrcMthDry = O
End Function

Function SrcMthFmToAy(Src$(), MthNm) As FmTo()
Dim FmAy&(): FmAy = SrcMthIdxAy(Src, MthNm)
Dim O() As FmTo, J%
For J = 0 To UB(FmAy)
    FmToPush O, FmTo(FmAy(J), SrcMthEndIdx(Src, FmAy(J)))
Next
SrcMthFmToAy = O
End Function

Function MdRplLin(Lno&, NewLin$, Optional A As CodeModule)
With DftMd(A)
    .DeleteLines Lno
    .InsertLines Lno, NewLin
End With
End Function
Sub SrcLinAssertIsMth(SrcLin)
If Not SrcLinIsMth(SrcLin) Then Stop
End Sub
Function SrcLinIsMth(SrcLin) As Boolean
SrcLinIsMth = SrcLinMthNm(SrcLin) <> ""
End Function
Function SrcLinRplMthMdy$(SrcLin, ToMdy)
SrcLinAssertIsMth SrcLin
SrcLinRplMthMdy = StrAppSpc(ToMdy) & SrcLinRmvMdy(SrcLin)
End Function
Function SrcLinRmvMdy$(SrcLin)
Dim O$: O = SrcLin
ParseMdy O
SrcLinRmvMdy = O
End Function
Function MdMthLnoAy(MthNm, Optional A As CodeModule)
MdMthLnoAy = AyIncN(SrcMthIdxAy(MdSrc(A), MthNm))
End Function
Function SrcMthIdxAyAll(Src$()) As Long()
If AyIsEmpty(Src) Then Exit Function
Dim O&(), J%
For J = 0 To UB(Src)
    If SrcLinMthTy(Src(J)) <> "" Then Push O, J
Next
SrcMthIdxAyAll = O
End Function
Private Function SrcMthIdxAy_%(Src$(), MthNm$, Optional FmIdx%)
If AyIsEmpty(Src) Then SrcMthIdxAy_% = -1: Exit Function
Dim MthIdx&
Dim M As SrcLinBrk
For MthIdx = FmIdx To UB(Src)
    M = SrcLinBrk(Src(MthIdx))
    If MthNm <> M.MthNm Then GoTo X
    SrcMthIdxAy_ = MthIdx
    Exit Function
X:
Next
SrcMthIdxAy_ = -1
End Function
Function SrcMthIdxFm%(Src$(), MthNm, Optional Fm&)
Dim I%
    For I = 0 To UB(Src)
        If SrcLinIsMth(Src(I)) Then SrcMthIdxFm% = I: Exit Function
    Next
SrcMthIdxFm = -1
End Function

Function SrcMthIdxAy(Src$(), MthNm) As Long()
If AyIsEmpty(Src) Then Exit Function
Dim I%
    I = SrcMthIdxFm(Src, MthNm)
If I = -1 Then Exit Function
Dim O&()
    Push O, I
    If SrcLinMthTy(Src(I)) = "Property" Then
        Dim J%, Fm&
        For J = 1 To 3
            Fm = I + 1
            I = SrcMthIdxFm(Src, MthNm, Fm)
            If I = -1 Then GoTo X
            Push O, I
        Next
        Never
    End If
X:
SrcMthIdxAy = O
End Function

Function SrcMthIdxAyFstRmkIdx%(Src$(), MthIdx&)
Dim J%
For J = MthIdx - 1 To 0 Step -1
    If SrcLinIsCd(Src(J)) Then SrcMthIdxAyFstRmkIdx = J + 1: Exit Function
Next
Never
End Function

Function SrcMthLinCnt%(Src$(), FunLno&)
Stop: Stop
End Function

Function SrcMthLines$(Src$(), MthNm$)
SrcMthLines = JnCrLf(SrcMthLy(Src, MthNm))
End Function

Function SrcMthLno%(Src$(), MthNm$, Optional PrpTy$)
If AyIsEmpty(Src) Then Exit Function
If PrpTy <> "" Then
    If Not AyHas(Array("Get Let Set"), PrpTy) Then Stop
End If
Dim FunTy$: FunTy = "Property " & PrpTy
Dim Lno&
Lno = 0
Const IMthNm% = 2
Dim M As SrcLinBrk
Dim Lin
For Each Lin In Src
    Lno = Lno + 1
    M = SrcLinBrk(Lin)
    If M.MthNm = "" Then GoTo Nxt
    If M.MthNm <> MthNm Then GoTo Nxt
    If PrpTy <> "" Then
        If M.Ty <> FunTy Then GoTo Nxt
    End If
    SrcMthLno = Lno
    Exit Function
Nxt:
Next
SrcMthLno = 0
End Function

Function SrcMthLy(Src$(), MthNm$) As String()
Dim A() As FmTo: A = SrcMthFmToAy(Src, MthNm)
Dim O$(), J%
For J = 0 To FmToUB(A)
    PushAy O, AySelFmTo(Src, A(J))
Next
SrcMthLy = O
End Function

Function SrcMthNy(Src$()) As String()
If AyIsEmpty(Src) Then Exit Function
Dim O$(), L
For Each L In Src
    PushNonEmpty O, SrcLinBrk(L).MthNm
Next
SrcMthNy = O
End Function
Function SrcLinMthNm$(SrcLin)
Dim L$: L = SrcLin
ParseMdy L
Dim MthTy$
    MthTy = ParseMthTy(L)
If MthTy = "" Then Exit Function
If MthTy = "Property" Then
    If ParsePrpTy(L) = "" Then Stop
End If
SrcLinMthNm = ParseNm(L)
End Function

Function SrcLinBrk(SrcLin) As SrcLinBrk
Dim L$: L = SrcLin
Dim O As SrcLinBrk
O.Mdy = ParseMdy(L)
O.Ty = ParseMthTy(L)
Select Case O.Ty
Case "Property": O.Ty = ParsePrpTy(L)
Case "": Exit Function
End Select
If O.Ty = "" Then Stop
O.MthNm = ParseNm(L)
If O.MthNm = "" Then Stop
SrcLinBrk = O
End Function

Function SrcLinMthTy$(SrcLin)
Dim L$: L = Trim(SrcLin)
ParseMdy L
SrcLinMthTy = ParseMthTy(L)
End Function

Function SrcLinIsEnm(SrcLin) As Boolean
Dim L$: L = SrcLin
ParseMdy L
SrcLinIsEnm = IsPfx(L, "Enum")
End Function

Function SrcLinIsRmk(SrcLin) As Boolean
SrcLinIsRmk = FstChr(LTrim(SrcLin)) = "'"
End Function

Function SrcLinIsTy(SrcLin) As Boolean
Dim L$: L = SrcLin
ParseMdy L
SrcLinIsTy = IsPfx(L, "Type")
End Function

Function SrcTyCnt%(Src$())
If AyIsEmpty(Src) Then Exit Function
Dim I, O%
For Each I In Src
    If SrcLinIsTy(I) Then O = O + 1
Next
SrcTyCnt = O
End Function

Function SyOfDfnTy() As String()
Static X As Boolean, Y
If Not X Then
    X = True
    Y = Sy("Function ", "Sub ", "Property ", "Type ", "Enum ")
End If
SyOfDfnTy = Y
End Function
Function SyOfMthTy() As String()
Static X As Boolean, Y
If Not X Then
    X = True
    Y = Sy("Function ", "Sub ", "Property ")
End If
SyOfMthTy = Y
End Function
Function SyOfPrpTy() As String()
Static X As Boolean, Y
If Not X Then
    X = True
    Y = Sy("Get ", "Set ", "Let ")
End If
SyOfPrpTy = Y
End Function

Function SyOfMdy() As String()
Static X As Boolean, Y
If Not X Then
    X = True
    Y = Sy("Public ", "Private ", "Friend ")
End If
SyOfMdy = Y
End Function

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

Private Function SrcMthEndIdx&(Src$(), MthIdx&)
Dim MthLin$
    MthLin = Src(MthIdx)
    
Dim Pfx$
    Pfx = SrcLinEndLinPfx(MthLin)
Dim O&
    For O = MthIdx + 1 To UB(Src)
        If IsPfx(Src(O), Pfx) Then SrcMthEndIdx = O: Exit Function
    Next
Er "SrcMthEndIdx: In {Src} {MthIdx} has {MthLin}, cannot find {FunEndLinPfx} in lines after [MthIdx]", Src, MthIdx, MthLin, Pfx
End Function

Private Function SrcMthIdxBdyLy(Src$(), MthIdx&) As String()
Dim ToIdx%: ToIdx = SrcMthEndIdx(Src, MthIdx)
Dim FmTo As FmTo
With FmTo
    .FmIdx = MthIdx
    .ToIdx = ToIdx
End With
Dim O$()
    O = AySelFmTo(Src, FmTo)
SrcMthIdxBdyLy = O
If AyLasEle(O) = "" Then Stop
End Function

Private Sub EnmBdyLy__Tst()
AyDmp EnmBdyLy("AA")
End Sub

Private Sub EnmLno__Tst()
Debug.Assert EnmLinIdx("AA", Md("Ide")) = 2
End Sub

Private Sub EnmMbrCnt__Tst()
Debug.Assert EnmMbrCnt("AA", Md("Ide")) = 1
End Sub

Private Sub JnContinueLin__Tst()
Dim O$(3)
O(0) = "A _"
O(1) = "B _"
O(2) = "C"
O(3) = "D"
Dim Act$(): Act = JnContinueLin(O)
Debug.Assert UB(Act) = 3
Debug.Assert Act(0) = "A B C"
Debug.Assert Act(1) = ""
Debug.Assert Act(2) = ""
Debug.Assert Act(3) = "D"
End Sub

Private Sub MdAppLines__Tst()
Const MdNm$ = "Module1"
Dim M As CodeModule
    Set M = Md(MdNm)
MdAppLines "'aa", M
End Sub

Private Sub MdMthDrs__Tst()
DrsBrw MdMthDrs(WithBdyLy:=True, A:=Md("Vb_Str"))
End Sub

Private Sub MdMthLines__Tst()
Debug.Print Len(MdMthLines("MdMthLines"))
Debug.Print MdMthLines("MdMthLines")
End Sub

Function MdInfDt__Tst()
DtBrw MdInfDt
End Function

Private Sub MdLy__Tst()
AyBrw MdLy
End Sub

Private Sub MdSrt__Tst()
MdSrt Md("bb_Lib_Acs")
End Sub

Private Sub MdSrtedBdyLines__Tst()
StrBrw MdSrtedBdyLines(Md("Vb_Str"))
End Sub

Private Sub MdTstSubLines__Tst()
Dim Ny$()
    Ny = PjMdNy
    Ny = AySrt(Ny)
    
Dim I, M As CodeModule

Dim Dr(), Dry()
For Each I In Ny
    Set M = Md(I)
    Dr = Array(I, MdTstSubLines(M))
    Push Dry, Dr
Next
Dim Drs As Drs
    Drs.Dry = Dry
    Drs.Fny = SplitSpc("Md TstSub")
    Drs = DrsExpLinesCol(Drs, "TstSub")
DrsBrw Drs, , "Md"
End Sub

Private Sub PjMthDrs__Tst()
Dim Drs As Drs
Drs = PjMthDrs(WithBdyLines:=True)
WsVis DrsWs(Drs, PjNm)
End Sub

Private Sub PjMdAy__Tst()
Dim O() As CodeModule
O = PjMdAy
Dim I, Md As CodeModule
For Each I In O
    Set Md = I
    Debug.Print MdNm(Md)
Next
End Sub

Private Sub PjMdNy__Tst()
AyBrw PjMdNy
End Sub

Private Sub SrcDclCnt__Tst()
Dim Src$(): Src = MdSrc(Md("LnkT"))
Debug.Assert SrcDclCnt(Src) = 15
End Sub

Private Sub SrcFstMthIdx__Tst()
Dim Src$(): Src = MdSrc
Debug.Assert SrcFstMthIdx(Src) = 19
End Sub

Sub SrcMthDry__Tst()
Const MdNm$ = "DaoDb"
Dim Src$(): Src = MdSrc(Md(MdNm))
DryBrw SrcMthDry(Src, MdNm)
End Sub

Sub SrcMthIdxAyAll__Tst()
Dim Src$(): Src = MdSrc(Md("DaoDb"))
Dim Ay$(): Ay = AySelByIdxAy(Src, SrcMthIdxAyAll(Src))
AyBrw Ay
End Sub

Private Sub SrcLinBrk__Tst()
Dim Act As SrcLinBrk:
Act = SrcLinBrk("Private Function AA()")
Debug.Assert Act.Mdy = "Private"
Debug.Assert Act.Ty = "Function"
Debug.Assert Act.MthNm = "AA"

Act = SrcLinBrk("Private Sub TakBet__Tst()")
Debug.Assert Act.Mdy = "Private"
Debug.Assert Act.Ty = "Sub"
Debug.Assert Act.MthNm = "TakBet__Tst"
End Sub

Sub Tst()
EnmBdyLy__Tst
EnmLno__Tst
EnmMbrCnt__Tst
JnContinueLin__Tst
MdMthDrs__Tst
MdInfDt__Tst
MdLy__Tst
MdSrt__Tst
MdSrtedBdyLines__Tst
PjMthDrs__Tst
PjMdAy__Tst
PjMdNy__Tst
SrcDclCnt__Tst
SrcFstMthIdx__Tst
SrcMthDry__Tst
SrcLinBrk__Tst
End Sub

