Attribute VB_Name = "bb_Lib_Ide"
Option Compare Database
Option Explicit
Type LnoCnt
    Lno As Long
    Cnt As Long
End Type
Type SrcLinBrk
    FunNm As String
    Ty As String
    Mdy As String
End Type
Sub MdCpy(Fm$, ToMdNm$, Optional A As VBProject)
Dim FmMd As CodeModule: Set FmMd = Md(Fm, A)
Dim Ty As vbext_ComponentType: Ty = MdTy(FmMd)
MdNew ToMdNm, Ty, A
Dim O As CodeModule: Set O = Md(ToMdNm)
MdAppLy MdLy(FmMd), O
End Sub

Function MdTy(Optional A As CodeModule) As vbext_ComponentType
MdTy = DftMd(A).Parent.Type
End Function

Sub MdRmvFun(FunNm$, Optional A As CodeModule)
MdRmvLnoCnt MdFunLnoCnt(FunNm, A)
End Sub
Sub MdRmvLnoCnt(LnoCnt As LnoCnt, Optional A As CodeModule)
If LnoCnt.Cnt = 0 Then Exit Sub
DftMd(A).DeleteLines LnoCnt.Lno, LnoCnt.Cnt
End Sub
Function LnoCntToStr$(A As LnoCnt)
LnoCntToStr = FmtQQ("Lno(?) Cnt(?)", A.Lno, A.Cnt)
End Function
Private Sub MdFunLnoCnt__Tst()
Debug.Print LnoCntToStr(MdFunLnoCnt("MdFunLnoCnt"))
End Sub
Function MdFunLnoCnt(FunNm$, Optional A As CodeModule) As LnoCnt
Dim Md As CodeModule: Set Md = DftMd(A)
Dim Lno&
Lno = 0
Dim Lin
Const IFunNm% = 2
Dim M As SrcLinBrk
For Each Lin In MdLy(Md)
    Lno = Lno + 1
    M = SrcLinBrk(Lin)
    If M.FunNm <> "" Then
        If FunNm = M.FunNm Then
            Dim O As LnoCnt
            Dim EndLno&: EndLno = MdFunEndLno(Md, Lno, M.Ty)
            Dim Cnt&
            Cnt = EndLno - Lno + 1
            O.Lno = Lno
            O.Cnt = Cnt
            MdFunLnoCnt = O
            Exit Function
        End If
    End If
Next
End Function
Private Function MdFunEndLno&(A As CodeModule, BegLno&, Ty$)
Dim Pfx$
    If AyHas(SplitSpc("Get Let Set"), Ty) Then
        Pfx = "End Property"
    Else
        Pfx = "End " & Ty
    End If
Dim O&
    For O = BegLno + 1 To MdLasLno(A)
        If IsPfx(A.Lines(O, 1), Pfx) Then MdFunEndLno = O: Exit Function
    Next
Err.Raise 1, , FmtQQ("MdFunEndLno: No [?] in module[?] from Lno[?]", Pfx, MdNm(A), BegLno)
End Function

Private Function IsPfx(S, Pfx) As Boolean
IsPfx = (Left(S, Len(Pfx)) = Pfx)
End Function

Sub MdAppLy(Ly$(), Optional A As CodeModule)
If IsEmptyAy(Ly) Then Exit Sub
DftMd(A).InsertLines MdLasLno(A) + 1, JnCrLf(Ly)
End Sub

Function MdLasLno&(Optional A As CodeModule)
MdLasLno = DftMd(A).CountOfLines
End Function
Sub MdUpdTstFun(Optional A As CodeModule)
Dim Md As CodeModule
Set Md = DftMd(A)
Select Case MdTy(Md)
Case _
    vbext_ComponentType.vbext_ct_StdModule, _
    vbext_ComponentType.vbext_ct_MSForm, _
    vbext_ComponentType.vbext_ct_ClassModule, _
    vbext_ComponentType.vbext_ct_Document
    MdRmvFun "Tst", Md
    MdAppLy MdTstFunLy(Md), Md
End Select
End Sub

Sub PjUpdTstFun(Optional A As VBProject)
Dim I, Md As CodeModule
For Each I In PjMdAy(A)
    Set Md = I
    MdUpdTstFun Md
Next
End Sub

Sub PjExp(Optional A As VBProject)
PthClrFil PjSrcPth(A)
Dim Md As CodeModule, I
For Each I In PjMdAy(A)
    Set Md = I
    MdExp Md
Next
End Sub

Sub MdSrt(Optional A As CodeModule)
MdExp A
Dim Ly$(): Ly = MdSrtedBdyLy(A)
MdRmvBdy A
MdAppLy Ly, A
End Sub

Private Sub MdSrtedBdyLy__Tst()
AyBrw MdSrtedBdyLy
End Sub
Function FunKey$(Mdy$, Ty$, FunNm$)
Dim A1 As Byte
    If IsSfx(FunNm, "__Tst") Then
        A1 = 9
    ElseIf FunNm = "Tst" Then
        A1 = 8
    Else
        Select Case Mdy
        Case "Public", "": A1 = 1
        Case "Friend": A1 = 2
        Case "Private": A1 = 3
        Case Else: Stop
        End Select
    End If
FunKey = FmtQQ("?:?:?", A1, FunNm, Ty)
End Function
Function MdSrtedBdyLy(Optional A As CodeModule) As String()
Dim Drs As Drs: Drs = MdFunDrs(WithBdyLy:=True, A:=A)
Dim Ky$()
    Erase Ky
    Dim Ty$, Mdy$, FunNm$, K$, BdyLy$(), IdxAy&(), Dr
    IdxAy = FidxAy(Drs.Fny, "Mdy FunNm Ty BdyLy")
    For Each Dr In Drs.Dry
        AyAsg_Idx Dr, IdxAy, Mdy, FunNm, Ty, BdyLy
        Push Ky, FunKey(Mdy, Ty, FunNm)
        If Not IsPfx(AyLasEle(BdyLy), "End") Then Stop
    Next
Dim I&()
    AyBrw AySrt(Ky)
    Stop
    I = AySrtIntoIdxAy(Ky)
Dim O$(), J&
J = 0
Dim IBdyLy%: IBdyLy = AyIdx(Drs.Fny, "BdyLy")
For J = 0 To UB(I)
    Push O, ""
    Dr = Drs.Dry(I(J))
    BdyLy = Dr(IBdyLy)
    PushAy O, BdyLy
Next
MdSrtedBdyLy = O
If Sz(O) < MdBdyLinCnt(A) Then Stop
End Function
Function MdBdyLinCnt&(Optional A As CodeModule)
Dim M As CodeModule: Set M = DftMd(A)
MdBdyLinCnt = M.CountOfLines - M.CountOfDeclarationLines
End Function
Function MdTstFunNy(Optional A As CodeModule) As String()
Dim M As Drs: M = MdFunDrs(A:=A)
Dim Dr
Dim O$(), Mdy$, Ty$, FunNm$, IFunNm%
Fiy M.Fny, "FunNm", IFunNm
For Each Dr In M.Dry
    FunNm = Dr(IFunNm)
    If IsSfx(FunNm, "__Tst") Then
        Push O, FunNm
    End If
Next
MdTstFunNy = O
End Function
Function PjTstFunNy_WithEr(Optional A As VBProject) As String()
Dim O$(), I, Md As CodeModule
For Each I In PjMdAy(A)
    Set Md = I
    PushAy O, AddAyPfx(MdTstFunNy_WithEr(Md), MdNm(Md) & ".")
Next
PjTstFunNy_WithEr = O
End Function
Function MdTstFunNy_WithEr(Optional A As CodeModule) As String()
Dim M As Drs: M = MdFunDrs(A:=A)
Dim O$(), Mdy$, Ty$, FunNm$, Dr
Dim IMdy%, ITy%, IFunNm%
Fiy M.Fny, "Mdy Ty FunNm", IMdy, ITy, IFunNm
If IsEmptyAy(M.Dry) Then Exit Function
For Each Dr In M.Dry
    If IsSfx(Dr(IFunNm), "__Tst") Then
        If Dr(IMdy) <> "Private" Or Dr(ITy) <> "Sub" Then
            Push O, Dr(IFunNm)
        End If
    End If
Next
MdTstFunNy_WithEr = O
End Function

Private Sub MdFunDrs__Tst()
DrsBrw MdFunDrs(WithBdyLy:=True)
End Sub
Private Sub PjFunDrs__Tst()
Dim Drs As Drs
Drs = PjFunDrs(WithBdyLines:=True)
WsVis DrsWs(Drs, PjNm)
End Sub
Function PjNm$(Optional A As VBProject)
PjNm = DftPj(A).Name
End Function

Sub AAA()
PjFunDrs__Tst
End Sub

Function PjFunDrs(Optional WithBdyLy As Boolean, Optional WithBdyLines As Boolean, Optional A As VBProject) As Drs
Dim Dry()
    Dim I, Md As CodeModule
    For Each I In PjMdAy(A)
        Set Md = I
        PushAy Dry, MdFunDrs(WithBdyLy, WithBdyLines, A:=Md).Dry
    Next
Dim O As Drs
    O.Fny = MdFunDrsFny(WithBdyLy, WithBdyLines)
    O.Dry = Dry
PjFunDrs = O
End Function

Function IsEmptyMd(Optional A As CodeModule) As Boolean
IsEmptyMd = DftMd(A).CountOfLines = 0
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
Function MdFunDrsFny(WithBdyLy As Boolean, WithBdyLines As Boolean) As String()
Dim O$(): O = SplitSpc("Lno Mdy Ty FunNm MdNm")
If WithBdyLy Then Push O, "BdyLy"
If WithBdyLines Then Push O, "BdyLines"
MdFunDrsFny = O
End Function
Function MdFunDrs(Optional WithBdyLy As Boolean, Optional WithBdyLines As Boolean, Optional A As CodeModule) As Drs
Dim Dry()
    Dim Md As CodeModule: Set Md = DftMd(A)
    Dim M As SrcLinBrk
    If Not IsEmptyMd(A) Then
        Dim Dr(), Lin
        Dim MNm$: MNm = MdNm(Md)
        Dim Lno&, J&
        For Lno = Md.CountOfDeclarationLines + 1 To Md.CountOfLines
            Lin = MdContLin(Lno, Md)
            M = SrcLinBrk(Lin)
            If M.FunNm <> "" Then
                Dr = Array(Lno, M.Mdy, M.Ty, M.FunNm, MNm)
                If WithBdyLy Then Push Dr, MdFunDrs_FunBdyLy(Md, Lno, M.Ty)
                If WithBdyLines Then Push Dr, JnCrLf(MdFunDrs_FunBdyLy(Md, Lno, M.Ty))
                Push Dry, Dr
            End If
        Next
    End If
Dim O As Drs
    O.Fny = MdFunDrsFny(WithBdyLy, WithBdyLines)
    O.Dry = Dry
MdFunDrs = O
End Function
Private Function MdFunDrs_FunBdyLy(A As CodeModule, Lno&, Ty$) As String()
Dim EndLno&: EndLno = MdFunEndLno(A, Lno, Ty)
Dim Cnt&: Cnt = EndLno - Lno + 1
MdFunDrs_FunBdyLy = SplitCrLf(A.Lines(Lno, Cnt))
End Function
Private Sub SrcLinBrk__Tst()
Dim Act As SrcLinBrk: Act = SrcLinBrk("Private Function AA()")
Debug.Assert Act.Mdy = "Private"
Debug.Assert Act.Ty = "Function"
Debug.Assert Act.FunNm = "AA"
End Sub
Function SrcLinBrk(SrcLin) As SrcLinBrk
Dim L$: L = SrcLin
Dim O As SrcLinBrk
O.Mdy = ParseMdy(L)
O.Ty = ParseFunTy(L): O.Ty = RmvPfx(O.Ty, "Property ")
If O.Ty <> "" Then O.FunNm = ParseNm(L)
SrcLinBrk = O
End Function
Function ParseMdy$(OLin$)
ParseMdy = ParseOneOf(OLin, Sy("Public", "Private", "Friend"))
End Function
Function ParseFunTy$(OLin$)
ParseFunTy = ParseOneOf(OLin, Sy("Function", "Sub", "Property Get", "Property Let", "Property Set", "Type", "Enum"))
End Function
Function ParseOneOf(OLin$, OneOfAy$())
Dim I
For Each I In OneOfAy
    If IsPfx(OLin, I) Then OLin = RmvFstNChr(RmvPfx(OLin, I)): ParseOneOf = I: Exit Function
Next
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
Sub MdRmvBdy(Optional A As CodeModule)
MdRmvLnoCnt MdBdyLnoCnt(A)
End Sub
Function MdBdyLnoCnt(Optional A As CodeModule) As LnoCnt
Dim Md As CodeModule: Set Md = DftMd(A)
Dim O As LnoCnt
O.Lno = Md.CountOfDeclarationLines + 1
O.Cnt = Md.CountOfLines - O.Lno
MdBdyLnoCnt = O
End Function
Function MdTstFunLy(Optional A As CodeModule) As String()
Dim O$(), Ay$()
Ay = MdTstFunNy(A)
If IsEmptyAy(Ay) Then Exit Function
Push O, "Sub Tst()"
PushAy O, AySrt(MdTstFunNy(A))
Push O, "End Sub"
MdTstFunLy = O
End Function
Function DftPj(Optional A As VBProject) As VBProject
If IsNothing(A) Then
    Set DftPj = Application.VBE.ActiveVBProject
Else
    Set DftPj = A
End If
End Function
Private Sub MdSrt__Tst()
MdSrt Md("bb_Lib_Acs")
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
Function PjMdAy(Optional A As VBProject) As CodeModule()
Dim O() As CodeModule
Dim Cmp As VBComponent
For Each Cmp In DftPj(A).VBComponents
    PushObj O, Cmp.CodeModule
Next
PjMdAy = O
End Function
Sub BrwPjSrc()
PthBrw PjSrcPth
End Sub
Private Sub MdLy__Tst()
AyBrw MdLy
End Sub
Function MdLy_Jn(Optional A As CodeModule) As String()
MdLy_Jn = JnContinueLin(MdLy(A))
End Function
Function MdLy(Optional A As CodeModule) As String()
Dim Md As CodeModule: Set Md = DftMd(A)
Dim N&: N = Md.CountOfLines
If N = 0 Then Exit Function
Dim O$()
ReDim O(N - 1)
Dim J&
For J = 0 To N - 1
    O(J) = Md.Lines(J + 1, 1)
Next
MdLy = O
End Function
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
Sub PjSrcPthBrw(Optional A As VBProject)
PthBrw PjSrcPth(A)
End Sub
Function PjSrcPth$(Optional A As VBProject)
Dim Ffn$: Ffn = DftPj(A).FileName
Dim Fn$: Fn = FfnFn(Ffn)
Dim O$:
O = FfnPth(DftPj(A).FileName) & "Src\": PthEns O
O = O & Fn & "\":                       PthEns O
PjSrcPth = O
End Function
Function MdSrcExt$(Optional A As CodeModule)
Dim O$
Select Case MdCmpTy(A)
Case vbext_ct_ClassModule: O = ".cls"
Case vbext_ct_Document: O = ".cls"
Case vbext_ct_StdModule: O = ".bas"
Case Else: Err.Raise 1, , "MdSrcExt: Unexpected MdCmpTy.  Should be [Class or Module or Document]"
End Select
MdSrcExt = O
End Function
Function MdCmpTy(Optional A As CodeModule) As vbext_ComponentType
MdCmpTy = MdCmp(A).Type
End Function
Function MdCmp(Optional A As CodeModule) As VBComponent
Set MdCmp = DftMd(A).Parent
End Function

Function MdSrcFn$(Optional A As CodeModule)
MdSrcFn = MdCmp(A).Name & MdSrcExt(A)
End Function
Function MdSrcFfn$(Optional A As CodeModule)
MdSrcFfn = PjSrcPth(MdPj(A)) & MdSrcFn(A)
End Function
Function MdPj(Optional A As CodeModule) As VBProject
Set MdPj = DftMd(A).Parent.Collection.Parent
End Function
Function MdExp(Optional A As CodeModule)
Dim Md As CodeModule: Set Md = DftMd(A)
Md.Parent.Export MdSrcFfn(Md)
Debug.Print MdNm(A)
End Function
Function MdNm(Optional A As CodeModule)
MdNm = DftMd(A).Parent.Name
End Function
Function Pj(PjNm) As VBProject
Set Pj = Application.VBE.VBProjects(PjNm)
End Function
Function DftMd(Optional A As CodeModule) As CodeModule
If IsNothing(A) Then
    Set DftMd = Application.VBE.ActiveCodePane.CodeModule
Else
    Set DftMd = A
End If
End Function
Function Md(MdNm$, Optional A As VBProject) As CodeModule
Set Md = DftPj(A).VBComponents(MdNm).CodeModule
End Function
Sub MdNew(Optional MdNm$, Optional Ty As vbext_ComponentType = vbext_ct_StdModule, Optional A As VBProject)
Dim O As VBComponent: Set O = DftPj(A).VBComponents.Add(Ty)
O.CodeModule.DeleteLines 1, 2
If MdNm <> "" Then O.Name = MdNm
End Sub

Sub MdRmv(MdNm$)
Dim C As VBComponent: Set C = Md(MdNm).Parent
C.Collection.Remove C
End Sub
Sub Tst()
JnContinueLin__Tst
MdFunDrs__Tst
MdFunLnoCnt__Tst
MdLy__Tst
MdSrtedBdyLy__Tst
MdSrt__Tst
PjFunDrs__Tst
PjMdAy__Tst
SrcLinBrk__Tst
End Sub
