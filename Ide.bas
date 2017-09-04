Attribute VB_Name = "Ide"
Option Explicit
Option Compare Database

Private Enum AA 'For testing
    AA1
    '
    
End Enum

Type LnoCnt
    Lno As Long
    Cnt As Long
End Type

Type SrcLinBrk
    FunNm As String
    Ty As String
    Mdy As String
End Type

Function DftMd(Optional a As CodeModule) As CodeModule
If IsNothing(a) Then
    Set DftMd = Application.VBE.ActiveCodePane.CodeModule
Else
    Set DftMd = a
End If
End Function

Function DftMdNm$(Nm)
If Nm = "" Then
    DftMdNm = MdNm(DftMd)
Else
    DftMdNm = Nm
End If
End Function

Function DftPj(Optional a As VBProject) As VBProject
If IsNothing(a) Then
    Set DftPj = Application.VBE.ActiveVBProject
Else
    Set DftPj = a
End If
End Function

Function EnmBdyLy(EnmNm$, Optional a As CodeModule) As String()
Dim B%: B = EnmLinIdx(EnmNm, a): If B = -1 Then Exit Function
Dim O$(), Ly$(), J%
Ly = MdDclLy(a)
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

Function EnmLinIdx%(EnmNm$, Optional a As CodeModule)
Dim Ly$(): Ly = MdDclLy(a)
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

Function EnmMbrCnt%(EnmNm$, Optional a As CodeModule)
EnmMbrCnt = Sz(EnmMbrLy(EnmNm, a))
End Function

Function EnmMbrLy(EnmNm$, Optional a As CodeModule) As String()
Dim Ly$(), O$(), J%
Ly = EnmBdyLy(EnmNm, a)
For J = 1 To UB(Ly) - 1
    If EnmIsMbrLin(Ly(J)) Then Push O, Ly(J)
Next
EnmMbrLy = O
End Function

Function FunDrsFny(WithBdyLy As Boolean, WithBdyLines As Boolean) As String()
Dim O$(): O = SplitSpc("Lno Mdy Ty FunNm MdNm")
If WithBdyLy Then Push O, "BdyLy"
If WithBdyLines Then Push O, "BdyLines"
FunDrsFny = O
End Function

Function FunDrsKy(FunDrs As Drs) As String()
Dim Dry() As Variant: Dry = FunDrs.Dry
Dim Fny$(): Fny = FunDrs.Fny
Dim O$()
    Dim Ty$, Mdy$, FunNm$, K$, IdxAy&(), Dr
    IdxAy = FnyLIdxAy(Fny, "Mdy FunNm Ty")
    If AyIsEmpty(FunDrs.Dry) Then Exit Function
    For Each Dr In FunDrs.Dry
        'Debug.Print Mdy, FunNm, Ty
        AyAsg_Idx Dr, IdxAy, Mdy, FunNm, Ty
        Push O, FunKey(Mdy, Ty, FunNm)
    Next
FunDrsKy = O
End Function

Function FunKey$(Mdy$, Ty$, FunNm$)
Dim A1 As Byte
    If IsSfx(FunNm, "__Tst") Then
        A1 = 8
    ElseIf FunNm = "Tst" Then
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
FunKey = FmtQQ("?:?:?", A1, FunNm, A3)
End Function

Function FunLinEndLinPfx$(FunLin)
If FunLinIsPrp(FunLin) Then
    FunLinEndLinPfx = "End Property"
Else
    FunLinEndLinPfx = "End " & FunLinTy(FunLin)
End If
End Function

Function FunLinIsPrp(FunLin) As Boolean
Dim Ty$: Ty = FunLinTy(FunLin)
FunLinIsPrp = IsPfx(Ty, "Property")
End Function

Function FunLinTy$(FunLin)
Dim L$: L = Trim(FunLin)
ParseMdy L
Dim Ty$: Ty = ParseFunTy(L)
If Ty = "" Then Stop
FunLinTy = Ty
End Function

Function FunTyIsFun(FunTy$)
If FunTy = "" Then Exit Function
FunTyIsFun = True
If FunTy = "Function" Then Exit Function
If FunTy = "Sub" Then Exit Function
If FunTy = "Get" Then Exit Function
If FunTy = "Let" Then Exit Function
If FunTy = "SEt" Then Exit Function
FunTyIsFun = False
End Function

Function IsEmptyMd(Optional a As CodeModule) As Boolean
IsEmptyMd = DftMd(a).CountOfLines = 0
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

Function LinIsCd(Lin) As Boolean
Dim L$: L = Trim(Lin)
If L = "" Then Exit Function
If FstChr(L) = "'" Then Exit Function
LinIsCd = True
End Function

Function LinIsFunSubPrp(Lin) As Boolean
Dim L$: L = Lin
ParseMdy L
LinIsFunSubPrp = ParseFunSubPrp(L) <> ""
End Function

Function LnoCntStr$(a As LnoCnt)
LnoCntStr = FmtQQ("Lno(?) Cnt(?)", a.Lno, a.Cnt)
End Function

Function Md(Optional MdNm, Optional a As VBProject) As CodeModule
Set Md = DftPj(a).VBComponents(DftMdNm(MdNm)).CodeModule
End Function

Function MdAddOptExplicit(Optional a As CodeModule)
DftMd(a).InsertLines 1, "Option Explicit"
Debug.Print MdNm(a), "<-- Option Explicit added"
End Function

Sub MdAppLines(Lines$, Optional a As CodeModule)
If Lines = "" Then Exit Sub
DftMd(a).InsertLines MdLasLno(a) + 1, Lines
End Sub

Sub MdAppLy(Ly$(), Optional a As CodeModule)
MdAppLines JnCrLf(Ly), a
End Sub

Function MdBdyLines$(Optional a As CodeModule)
MdBdyLines = SrcBdyLines(MdSrc(a))
End Function

Function MdBdyLnoCnt(Optional a As CodeModule) As LnoCnt
MdBdyLnoCnt = SrcBdyLnoCnt(MdSrc(a))
End Function

Function MdBdyLy(Optional a As CodeModule) As String()
MdBdyLy = SrcBdyLy(MdSrc(a))
End Function

Function MdCanHasCd(Optional a As CodeModule) As Boolean
Select Case MdTy(a)
Case _
    vbext_ComponentType.vbext_ct_StdModule, _
    vbext_ComponentType.vbext_ct_ClassModule, _
    vbext_ComponentType.vbext_ct_Document, _
    vbext_ComponentType.vbext_ct_MSForm
    MdCanHasCd = True
End Select
End Function

Function MdCmp(Optional a As CodeModule) As VBComponent
Set MdCmp = DftMd(a).Parent
End Function

Function MdCmpTy(Optional a As CodeModule) As vbext_ComponentType
MdCmpTy = MdCmp(a).Type
End Function

Function MdContLin$(Lno&, Optional a As CodeModule)
Dim J&, L&, Md As CodeModule: Set Md = a
L = Lno
Dim O$: O = Md.Lines(L, 1)
While LasChr(O) = "_"
    L = L + 1
    O = RmvLasChr(O) & Md.Lines(L, 1)
Wend
MdContLin = O
End Function

Sub MdCpy(Fm$, ToMdNm$, Optional a As VBProject)
Dim FmMd As CodeModule: Set FmMd = Md(Fm, a)
Dim Ty As vbext_ComponentType: Ty = MdTy(FmMd)
MdNew ToMdNm, Ty, a
Dim O As CodeModule: Set O = Md(ToMdNm)
MdAppLy MdLy(FmMd), O
End Sub

Function MdDclLines$(Optional a As CodeModule)
With DftMd(a)
    If .CountOfDeclarationLines = 0 Then Exit Function
    MdDclLines = .Lines(1, .CountOfDeclarationLines)
End With
End Function

Function MdDclLy(Optional a As CodeModule) As String()
MdDclLy = SplitCrLf(MdDclLines(a))
End Function

Function MdEnmCnt%(Optional a As CodeModule)
MdEnmCnt = SrcEnmCnt(MdDclLy(a))
End Function

Sub MdEnsOptExplicit(Optional a As CodeModule)
If Not MdHasOptExplicit(a) Then MdAddOptExplicit a
'If MdHasOptExplicit(A) Then
'    Debug.Print MdNm(A), "(* With Option Explicit *)"
'Else
'    Debug.Print MdNm(A), "<-------------------- No Option Explicit"
'End If
End Sub

Function MdExp(Optional a As CodeModule)
Dim Md As CodeModule: Set Md = DftMd(a)
Md.Parent.Export MdSrcFfn(Md)
Debug.Print MdNm(a)
End Function

Function MdFunCnt%(Optional a As CodeModule)
MdFunCnt = SrcFunCnt(MdSrc(a))
End Function

Function MdFunDrs(Optional WithBdyLy As Boolean, Optional WithBdyLines As Boolean, Optional a As CodeModule) As Drs
MdFunDrs = SrcFunDrs(MdSrc(a), MdNm(a), WithBdyLy, WithBdyLines)
End Function

Function MdFunLines$(FunNm$, Optional PrpTy$, Optional a As CodeModule)
MdFunLines = SrcFunLines(MdSrc(a), FunNm, PrpTy)
End Function

Function MdFunLnoCnt(FunNm$, Optional PrpTy$, Optional a As CodeModule) As LnoCnt
Stop
End Function

Function MdFunNy(Optional a As CodeModule) As String()
MdFunNy = SrcFunNy(MdSrc(a))
End Function

Function MdHasOptExplicit(Optional a As CodeModule)
Dim Ay$()
    Ay = MdDclLy(a)
Dim I
For Each I In Ay
    If I = "Option Explicit" Then MdHasOptExplicit = True: Exit Function
Next
End Function

Function MdInfDr(Optional a As CodeModule) As Variant()
Dim Pj$, Md$, IsCls As Boolean, LinCnt%, FunCnt%, TyCnt%, EnmCnt%
Dim M As CodeModule: Set M = DftMd(a)
Pj = PjNm(MdPj(M))
Md = MdNm(M)
LinCnt = M.CountOfLines
FunCnt = MdFunCnt(M)
TyCnt = MdTyCnt(M)
EnmCnt = MdEnmCnt(M)
IsCls = MdIsCls(M)
MdInfDr = Array(Pj, Md, IsCls, LinCnt, FunCnt, TyCnt, EnmCnt)
End Function

Function MdInfDry(Optional a As VBProject) As Variant()
Dim I, M As CodeModule, O()
For Each I In PjMdAy(a)
    Set M = I
    Push O, MdInfDr(M)
Next
MdInfDry = O
End Function

Function MdInfDt(Optional a As VBProject) As Dt
Dim O As Dt
O.DtNm = "MdInf"
O.Fny = SplitSpc("Pj Md IsCls LinCnt FunCnt TyCnt EnmCnt")
O.Dry = MdInfDry(a)
MdInfDt = O
End Function

Function MdIsCls(Optional a As CodeModule) As Boolean
MdIsCls = MdTy(a) = vbext_ct_ClassModule
End Function

Function MdIsEmpty(Optional a As CodeModule)
MdIsEmpty = (DftMd(a).CountOfLines = 0)
End Function

Function MdIsExist(MdNm$, Optional a As VBProject) As Boolean
On Error GoTo X
MdIsExist = DftPj(a).VBComponents(MdNm).Name = MdNm
Exit Function
X:
End Function

Function MdLasLno&(Optional a As CodeModule)
MdLasLno = DftMd(a).CountOfLines
End Function

Function MdLines$(Optional a As CodeModule)
With DftMd(a)
    If .CountOfLines = 0 Then Exit Function
    MdLines = .Lines(1, .CountOfLines)
End With
End Function

Function MdLnoCntLines$(LnoCnt As LnoCnt, Optional a As CodeModule)
With LnoCnt
    If .Cnt = 0 Then Exit Function
    MdLnoCntLines = DftMd(a).Lines(.Lno, .Cnt)
End With
End Function

Function MdLy(Optional a As CodeModule) As String()
MdLy = SplitCrLf(MdLines(a))
End Function

Function MdLy_Jn(Optional a As CodeModule) As String()
MdLy_Jn = JnContinueLin(MdLy(a))
End Function

Sub MdNew(Optional MdNm$, Optional Ty As vbext_ComponentType = vbext_ct_StdModule, Optional a As VBProject)
Dim O As VBComponent: Set O = DftPj(a).VBComponents.Add(Ty)
O.CodeModule.DeleteLines 1, 2
If MdNm <> "" Then O.Name = MdNm
End Sub

Function MdNm(Optional a As CodeModule)
MdNm = DftMd(a).Parent.Name
End Function

Function MdPj(Optional a As CodeModule) As VBProject
Set MdPj = DftMd(a).Parent.Collection.Parent
End Function

Sub MdRmv(MdNm$)
Dim C As VBComponent: Set C = Md(MdNm).Parent
C.Collection.Remove C
End Sub

Sub MdRmvBdy(a As CodeModule)
MdRmvLnoCnt MdBdyLnoCnt(a), a
End Sub

Sub MdRmvFun(FunNm$, Optional PrpTy$, Optional a As CodeModule)
Dim M As LnoCnt: M = MdFunLnoCnt(FunNm, PrpTy$, a)
If M.Cnt = 0 Then
    Debug.Print FmtQQ("Fun[?] in Md[?] not found, cannot Rmv", FunNm, MdNm(a))
Else
    Debug.Print FmtQQ("Fun[?] in Md[?] is removed", FunNm, MdNm(a))
End If
MdRmvLnoCnt M, a
End Sub

Sub MdRmvLnoCnt(LnoCnt As LnoCnt, a As CodeModule)
With LnoCnt
    If .Cnt = 0 Then Exit Sub
    DftMd(a).DeleteLines .Lno, .Cnt
End With
End Sub

Sub MdRmvTstFun(Optional a As CodeModule)
MdRmvFun "Tst", , a
End Sub

Function MdSrc(Optional a As CodeModule) As String()
MdSrc = MdLy(a)
End Function

Function MdSrcFfn$(Optional a As CodeModule)
MdSrcFfn = PjSrcPth(MdPj(a)) & MdSrcFn(a)
End Function

Function MdSrcFn$(Optional a As CodeModule)
MdSrcFn = MdCmp(a).Name & MdSrcExt(a)
End Function

Sub MdSrt(Optional a As CodeModule)
Dim Md As CodeModule: Set Md = DftMd(a)
Debug.Print MdNm(Md),
Dim Old$: Old = MdBdyLines(Md)
Dim Lines$: Lines = MdSrtedBdyLines(Md)
If Old = Lines Then
    Debug.Print "<=== Same"
    Exit Sub
End If
Debug.Print
MdRmvBdy Md
Debug.Print "<<-- Rmv"
MdAppLines Lines, Md
Debug.Print "<<-- Ins"
End Sub

Function MdSrtedBdyLines$(Optional a As CodeModule)
If MdIsEmpty(a) Then Exit Function
Dim Drs As Drs: Drs = MdFunDrs(WithBdyLines:=True, a:=a)
Dim Ky$(): Ky = FunDrsKy(Drs)
Dim I&()
    I = AySrtIntoIdxAy(Ky)
Dim O$()
Dim J%
    Dim IBdyLines%: IBdyLines = AyIdx(Drs.Fny, "BdyLines")
    Dim Dr
    For J = 0 To UB(I)
        Dr = Drs.Dry()(I(J))
        Push O, vbCrLf & Dr(IBdyLines)
    Next
MdSrtedBdyLines = JnCrLf(O)
End Function

Function MdTstFunLines$(Optional a As CodeModule)
MdTstFunLines = JnCrLf(MdTstFunLy(a))
End Function

Function MdTstFunLy(Optional a As CodeModule) As String()
Dim O$(), Ay$()
Ay = MdTstFunNy(a)
If AyIsEmpty(Ay) Then Exit Function
Push O, "Sub Tst()"
PushAy O, AySrt(MdTstFunNy(a))
Push O, "End Sub"
MdTstFunLy = O
End Function

Function MdTstFunNy(Optional a As CodeModule) As String()
If MdIsEmpty(a) Then Exit Function
Dim M As Drs: M = MdFunDrs(a:=a)
Dim Dr
Dim O$(), Mdy$, Ty$, FunNm$, IFunNm%
Fiy M.Fny, "FunNm", IFunNm
If AyIsEmpty(M.Dry) Then Exit Function
For Each Dr In M.Dry
    FunNm = Dr(IFunNm)
    If IsSfx(FunNm, "__Tst") Then
        Push O, FunNm
    End If
Next
MdTstFunNy = O
End Function

Function MdTstFunNy_WithEr(Optional a As CodeModule) As String()
Dim M As Drs: M = MdFunDrs(a:=a)
Dim O$(), Mdy$, Ty$, FunNm$, Dr
Dim IMdy%, ITy%, IFunNm%
Fiy M.Fny, "Mdy Ty FunNm", IMdy, ITy, IFunNm
If AyIsEmpty(M.Dry) Then Exit Function
For Each Dr In M.Dry
    If IsSfx(Dr(IFunNm), "__Tst") Then
        If Dr(IMdy) <> "Private" Or Dr(ITy) <> "Sub" Then
            Push O, Dr(IFunNm)
        End If
    End If
Next
MdTstFunNy_WithEr = O
End Function

Function MdTy(Optional a As CodeModule) As vbext_ComponentType
MdTy = DftMd(a).Parent.Type
End Function

Function MdTyCnt%(Optional a As CodeModule)
MdTyCnt = SrcTyCnt(MdDclLy(a))
End Function

Sub MdUpdTstFun(Optional a As CodeModule)
If Not MdCanHasCd(a) Then Exit Sub
    
Dim NewLines$: NewLines = MdTstFunLines(a)
Dim OldLines$: OldLines = MdFunLines("Tst", , a)
If OldLines = NewLines Then
    Debug.Print FmtQQ("Fun[Tst] in Md[?] is same: [?] lines", MdNm(a), LinesLinCnt(OldLines))
    Exit Sub
End If

MdRmvTstFun a
Dim Ly$()
    Ly = MdTstFunLy(a)

If Sz(Ly) > 0 Then
    Debug.Print FmtQQ("Fun[Tst] in Md[?] is inserted", MdNm(a))
    MdAppLy Ly, a
End If
End Sub

Function ParseFunSubPrp$(OLin$)
ParseFunSubPrp = ParseOneOf(OLin, SyOfFunSubPrp)
End Function

Function ParseFunTy$(OLin$)
ParseFunTy = RTrim(ParseOneOf(OLin, SyOfFunTy))
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

Sub PjAssertNotUnderSrc(Optional a As VBProject)
Dim B$: B = PjPth(a)
If PthFdr(B) = "Src" Then Stop
End Sub

Sub PjCpyToSrc(Optional a As VBProject)
FilCpyToPth DftPj(a).FileName, PjSrcPth(a), OvrWrt:=True
End Sub

Sub PjEnsOptExplicit(Optional a As VBProject)
Dim I, Md As CodeModule
For Each I In PjMdAy(a)
    Set Md = I
    MdEnsOptExplicit Md
Next
End Sub

Sub PjExp(Optional a As VBProject)
PjAssertNotUnderSrc
PjCpyToSrc a
PthClrFil PjSrcPth(a)
Dim Md As CodeModule, I
For Each I In PjMdAy(a)
    Set Md = I
    MdExp Md
Next
End Sub

Function PjFunDrs(Optional WithBdyLy As Boolean, Optional WithBdyLines As Boolean, Optional a As VBProject) As Drs
Dim Dry()
    Dim I, Md As CodeModule
    For Each I In PjMdAy(a)
        Set Md = I
        PushAy Dry, MdFunDrs(WithBdyLy, WithBdyLines, a:=Md).Dry
    Next
Dim O As Drs
    O.Fny = FunDrsFny(WithBdyLy, WithBdyLines)
    O.Dry = Dry
PjFunDrs = O
End Function

Function PjMdAy(Optional a As VBProject) As CodeModule()
Dim O() As CodeModule
Dim Cmp As VBComponent
For Each Cmp In DftPj(a).VBComponents
    PushObj O, Cmp.CodeModule
Next
PjMdAy = O
End Function

Function PjMdNy(Optional a As VBProject) As String()
PjMdNy = OyPrp(PjMdAy(a), "Name", EmptySy)
End Function

Function PjNm$(Optional a As VBProject)
PjNm = DftPj(a).Name
End Function

Function PjPth$(Optional a As VBProject)
PjPth = FfnPth(DftPj(a).FileName)
End Function

Sub PjSrcBrw()
PthBrw PjSrcPth
End Sub

Function PjSrcPth$(Optional a As VBProject)
Dim Ffn$: Ffn = DftPj(a).FileName
Dim Fn$: Fn = FfnFn(Ffn)
Dim O$:
O = FfnPth(DftPj(a).FileName) & "Src\": PthEns O
O = O & Fn & "\":                       PthEns O
PjSrcPth = O
End Function

Sub PjSrt(Optional a As VBProject)
Dim Md As CodeModule, I
For Each I In PjMdAy(a)
    Set Md = I
    MdSrt Md
Next
End Sub

Function PjTstFunNy_WithEr(Optional a As VBProject) As String()
Dim O$(), I, Md As CodeModule
For Each I In PjMdAy(a)
    Set Md = I
    PushAy O, AyAddPfx(MdTstFunNy_WithEr(Md), MdNm(Md) & ".")
Next
PjTstFunNy_WithEr = O
End Function

Sub PjUpdTstFun(Optional a As VBProject)
Dim I, Md As CodeModule
For Each I In PjMdAy(a)
    Set Md = I
    MdUpdTstFun Md
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

Function SrcDclCnt%(Src$())
Dim I%: I = SrcFstFunIdx(Src)
If I = -1 Then SrcDclCnt = Sz(Src): Exit Function
SrcDclCnt = SrcFunIdxFstRmkIdx(Src, I)
End Function

Function SrcEnmCnt%(Src$())
If AyIsEmpty(Src) Then Exit Function
Dim I, O%
For Each I In Src
    If SrcLinIsEnm(I) Then O = O + 1
Next
SrcEnmCnt = O
End Function

Function SrcFstFunIdx%(Src$())
Dim J%
For J = 0 To UB(Src)
    If LinIsFunSubPrp(Src(J)) Then
        SrcFstFunIdx = J
        Exit Function
    End If
Next
SrcFstFunIdx = -1
End Function

Function SrcFunCnt%(Src$())
If AyIsEmpty(Src) Then Exit Function
Dim I, O%
For Each I In Src
    If SrcLinIsFun(I) Then O = O + 1
Next
SrcFunCnt = O
End Function

Function SrcFunDrs(Src$(), MdNm$, Optional WithBdyLy As Boolean, Optional WithBdyLines As Boolean) As Drs
SrcFunDrs.Dry = SrcFunDry(Src, MdNm$, WithBdyLy, WithBdyLines)
SrcFunDrs.Fny = FunDrsFny(WithBdyLy, WithBdyLines)
End Function

Function SrcFunDry(Src$(), MdNm$, Optional WithBdyLy As Boolean, Optional WithBdyLines As Boolean) As Variant()
Dim FunIdxAy%(): FunIdxAy = SrcFunIdxAy(Src)
Dim O()
    Dim M As SrcLinBrk
    Dim Dr(), Lno%
    Dim J&, FunLin$, LinIdx%
    Dim BdyLy$()
    For J = 0 To UB(FunIdxAy)
'        Debug.Print J
        LinIdx = FunIdxAy(J)
        Lno = LinIdx + 1
        FunLin = Src(LinIdx)
        M = SrcLinBrk(FunLin)
        Dr = Array(Lno, M.Mdy, M.Ty, M.FunNm, MdNm)
        If WithBdyLy Or WithBdyLines Then
            BdyLy = SrcFunIdxBdyLy(Src, LinIdx)
        End If
        If WithBdyLy Then Push Dr, BdyLy
        If WithBdyLines Then Push Dr, JnCrLf(BdyLy)
        Push O, Dr
    Next
SrcFunDry = O
End Function

Function SrcFunFmTo(Src$(), FunNm$, Optional PrpTy$) As FmTo
Dim Fm%: Fm = SrcFunIdx(Src, FunNm, PrpTy)
SrcFunFmTo.FmIdx = Fm
SrcFunFmTo.ToIdx = SrcFunIdxEndIdx(Src, Fm)
End Function

Function SrcFunIdx%(Src$(), FunNm$, Optional PrpTy$)
If AyIsEmpty(Src) Then SrcFunIdx% = -1: Exit Function
Dim FunIdx%
Dim M As SrcLinBrk
For FunIdx = 0 To UB(Src)
    M = SrcLinBrk(Src(FunIdx))
    If Not FunTyIsFun(M.Ty) Then GoTo X
    If FunNm <> M.FunNm Then GoTo X
    If PrpTy <> "" Then
        If M.Ty <> PrpTy Then GoTo X
    End If
    SrcFunIdx = FunIdx
    Exit Function
X:
Next
SrcFunIdx = -1
End Function

Function SrcFunIdxAy(Src$()) As Integer()
Dim O%()
    Dim Lin$
    Dim J&, Ty$
    For J = 0 To UB(Src)
        Ty = SrcLinBrk(Src(J)).Ty
        If FunTyIsFun(Ty) Then
            Push O, J
'            Debug.Print Ty, Src(J)
        End If
    Next
SrcFunIdxAy = O
End Function

Function SrcFunIdxFstRmkIdx%(Src$(), FunIdx%)
Dim J%
For J = FunIdx - 1 To 0 Step -1
    If LinIsCd(Src(J)) Then SrcFunIdxFstRmkIdx = J + 1: Exit Function
Next
Stop
End Function

Function SrcFunLinCntByLinIdx%(Src, LinIdx%)
End Function

Function SrcFunLines$(Src$(), FunNm$, Optional PrpTy$)
SrcFunLines = JnCrLf(SrcFunLy(Src, FunNm, PrpTy))
End Function

Function SrcFunLno%(Src$(), FunNm$, Optional PrpTy$)
If AyIsEmpty(Src) Then Exit Function
If PrpTy <> "" Then
    If Not AyHas(Array("Get Let Set"), PrpTy) Then Stop
End If
Dim FunTy$: FunTy = "Property " & PrpTy
Dim Lno&
Lno = 0
Const IFunNm% = 2
Dim M As SrcLinBrk
Dim Lin
For Each Lin In Src
    Lno = Lno + 1
    M = SrcLinBrk(Lin)
    If M.FunNm = "" Then GoTo Nxt
    If M.FunNm <> FunNm Then GoTo Nxt
    If PrpTy <> "" Then
        If M.Ty <> FunTy Then GoTo Nxt
    End If
    SrcFunLno = Lno
    Exit Function
Nxt:
Next
SrcFunLno = 0
End Function

Function SrcFunLy(Src$(), FunNm$, Optional PrpTy$) As String()
SrcFunLy = AyFmTo(Src, SrcFunFmTo(Src, FunNm, PrpTy))
End Function

Function SrcFunNy(Src$()) As String()
If AyIsEmpty(Src) Then Exit Function
Dim O$(), L
For Each L In Src
    PushNonEmpty O, SrcLinBrk(L).FunNm
Next
SrcFunNy = O
End Function

Function SrcLinBrk(SrcLin) As SrcLinBrk
Dim L$: L = SrcLin
Dim O As SrcLinBrk
O.Mdy = ParseMdy(L)
O.Ty = ParseFunTy(L): O.Ty = RmvPfx(O.Ty, "Property ")
If O.Ty <> "" Then O.FunNm = ParseNm(L)
SrcLinBrk = O
End Function

Function SrcLinFunTy$(SrcLin)
Dim L$: L = Trim(SrcLin)
ParseMdy L
SrcLinFunTy = SrcLinFunTy(L)
End Function

Function SrcLinIsEnm(SrcLin) As Boolean
Dim L$: L = SrcLin
ParseMdy L
SrcLinIsEnm = IsPfx(L, "Enum")
End Function

Function SrcLinIsFun(SrcLin) As Boolean
Dim L$: L = SrcLin
Dim a$
ParseMdy L
a = ParseFunTy(L): If a = "" Then Exit Function
If IsOneOf(a, Array("Type", "Enum")) Then Exit Function
SrcLinIsFun = True
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

Function SyOfFunSubPrp() As String()
Static X As Boolean, Y
If Not X Then
    X = True
    Y = Sy("Function ", "Sub ", "Function ", "Property ")
End If
SyOfFunSubPrp = Y
End Function

Function SyOfFunTy() As String()
Static X As Boolean, Y
If Not X Then
    X = True
    Y = Sy("Function ", "Sub ", "Property Get ", "Property Let ", "Property Set ", "Type ", "Enum ")
End If
SyOfFunTy = Y
End Function

Function SyOfMdy() As String()
Static X As Boolean, Y
If Not X Then
    X = True
    Y = Sy("Public ", "Private ", "Friend ")
End If
SyOfMdy = Y
End Function

Private Function MdSrcExt$(Optional a As CodeModule)
Dim O$
Select Case MdCmpTy(a)
Case vbext_ct_ClassModule: O = ".cls"
Case vbext_ct_Document: O = ".cls"
Case vbext_ct_StdModule: O = ".bas"
Case Else: Err.Raise 1, , "MdSrcExt: Unexpected MdCmpTy.  Should be [Class or Module or Document]"
End Select
MdSrcExt = O
End Function

Private Function SrcFunIdxBdyLy(Src$(), FunIdx%) As String()
Dim ToIdx%: ToIdx = SrcFunIdxEndIdx(Src, FunIdx)
Dim FmTo As FmTo
With FmTo
    .FmIdx = FunIdx
    .ToIdx = ToIdx
End With
Dim O$()
    O = AyFmTo(Src, FmTo)
SrcFunIdxBdyLy = O
If AyLasEle(O) = "" Then Stop
End Function

Private Function SrcFunIdxEndIdx%(Src$(), FunIdx%)
Dim FunLin$
    FunLin = Src(FunIdx)
    
Dim Pfx$
    Pfx = FunLinEndLinPfx(FunLin)
Dim O&
    For O = FunIdx + 1 To UB(Src)
        If IsPfx(Src(O), Pfx) Then SrcFunIdxEndIdx = O: Exit Function
    Next
Er "SrcFunIdxEndIdx: In {Src} {FunIdx} has {FunLin}, cannot find {FunEndLinPfx} in lines after [FunIdx]", Src, FunIdx, FunLin, Pfx
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

Private Sub MdFunDrs__Tst()
DrsBrw MdFunDrs(WithBdyLy:=True, a:=Md("Vb_Str"))
End Sub

Private Sub MdFunLines__Tst()
Debug.Print Len(MdFunLines("MdFunLines"))
Debug.Print MdFunLines("MdFunLines")
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

Private Sub MdTstFunLines__Tst()
Dim Ny$()
    Ny = PjMdNy
    Ny = AySrt(Ny)
    
Dim I, M As CodeModule

Dim Dr(), Dry()
For Each I In Ny
    Set M = Md(I)
    Dr = Array(I, MdTstFunLines(M))
    Push Dry, Dr
Next
Dim Drs As Drs
    Drs.Dry = Dry
    Drs.Fny = SplitSpc("Md TstFun")
    Drs = DrsExpLinesCol(Drs, "TstFun")
DrsBrw Drs, , "Md"
End Sub
Private Sub MdAppLines__Tst()
Const MdNm$ = "Module1"
Dim M As CodeModule
    Set M = Md(MdNm)
MdAppLines "'aa", M
End Sub
Private Sub PjFunDrs__Tst()
Dim Drs As Drs
Drs = PjFunDrs(WithBdyLines:=True)
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

Private Sub SrcFstFunIdx__Tst()
Dim Src$(): Src = MdSrc
Debug.Assert SrcFstFunIdx(Src) = 19
End Sub

Sub SrcFunDry__Tst()
Const MdNm$ = "DaoDb"
Dim Src$(): Src = MdSrc(Md(MdNm))
DryBrw SrcFunDry(Src, MdNm)
End Sub

Sub SrcFunIdxAy__Tst()
Dim Src$(): Src = MdSrc(Md("DaoDb"))
Dim Ay$(): Ay = AySelByIdxAy(Src, SrcFunIdxAy(Src))
AyBrw Ay
End Sub

Private Sub SrcLinBrk__Tst()
Dim Act As SrcLinBrk:
Act = SrcLinBrk("Private Function AA()")
Debug.Assert Act.Mdy = "Private"
Debug.Assert Act.Ty = "Function"
Debug.Assert Act.FunNm = "AA"

Act = SrcLinBrk("Private Sub TakBet__Tst()")
Debug.Assert Act.Mdy = "Private"
Debug.Assert Act.Ty = "Sub"
Debug.Assert Act.FunNm = "TakBet__Tst"
End Sub

Sub Tst()
EnmBdyLy__Tst
EnmLno__Tst
EnmMbrCnt__Tst
JnContinueLin__Tst
MdFunDrs__Tst
MdInfDt__Tst
MdLy__Tst
MdSrt__Tst
MdSrtedBdyLines__Tst
PjFunDrs__Tst
PjMdAy__Tst
PjMdNy__Tst
SrcDclCnt__Tst
SrcFstFunIdx__Tst
SrcFunDry__Tst
SrcFunIdxAy__Tst
SrcLinBrk__Tst
End Sub

