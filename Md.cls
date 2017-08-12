VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Md"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Public Md As CodeModule

Function FunLno&(FunNm)
If IsEmpty Then Exit Function
Dim O&
Dim A$(): A = Ly
Dim J&
For J = 1 To Md.CountOfLines
    With SrcLin(J)
        If .Nm = "" Then GoTo Nxt
        If FunNm = .Nm Then FunLno = J + 1: Exit Function
    End With
Nxt:
Next
End Function
Function SrcLin(Lno&) As SrcLin
Set SrcLin = ccNew.SrcLin(Md.Lines(Lno, 1))
End Function
Function EnmMbrCnt%(EnmNm$)
EnmMbrCnt = Sz(EnmMbrLy(EnmNm))
End Function
Private Sub EnmMbrCnt__Tst()
Debug.Assert EnmMbrCnt("AA") = 1
End Sub
Private Sub EnmLno__Tst()
Debug.Assert EnmLinIdx("AA") = 2
End Sub
Private Sub EnmBdyLy__Tst()
AyDmp EnmBdyLy("AA")
End Sub
Function EnmLinIdx%(EnmNm$)
Dim O&, L$
For O = 1 To Md.CountOfDeclarationLines
    With SrcLin(O)
        If Not .IsEnmLin Then GoTo Nxt
        If .Nm = EnmNm Then EnmLinIdx = O - 1: Exit Function
    End With
Nxt:
Next
EnmLinIdx = -1
End Function

Function EnmBdyLy(EnmNm$) As String()
Dim B%: B = EnmLinIdx(EnmNm): If B = -1 Then Exit Function
Dim O$(), Ly$(), J%
Ly = DclLy
For J = B To UB(Ly)
    Push O, Ly(J)
    If IsPfx(Ly(J), "End Enum") Then EnmBdyLy = O: Exit Function
Next
Stop
End Function
Function DclLy() As String()
DclLy = SplitCrLf(DclLines)
End Function
Function DclLines$()
With Md
    If .CountOfDeclarationLines = 0 Then Exit Function
    DclLines = .Lines(1, .CountOfDeclarationLines)
End With
End Function
Function EnmMbrLy(EnmNm$) As String()
Dim Ly$(), O$(), J&
Ly = EnmBdyLy(EnmNm)
For J = 1 To UB(Ly) - 1
    If Not SrcLin(J).IsRmkLin Then Push O, Ly(J)
Next
EnmMbrLy = O
End Function
Sub BrwSrc()
PthBrw Pjx.SrcPth
End Sub
Function Pjx() As Pj
Set Pjx = ccNew.Pj(Pj)
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

Function LnoCntToStr$(A As LnoCnt)
LnoCntToStr = FmtQQ("Lno(?) Cnt(?)", A.Lno, A.Cnt)
End Function

Sub AppLines(Lines$)
If Lines = "" Then Exit Sub
Md.InsertLines LasLno + 1, Lines
End Sub

Sub AppLy(Ly$())
AppLines JnCrLf(Ly)
End Sub

Function BdyLnoCnt() As LnoCnt
Dim O As LnoCnt
O.Lno = Md.CountOfDeclarationLines + 1
O.Cnt = Md.CountOfLines - O.Lno + 1
BdyLnoCnt = O
End Function

Function Cmp() As VBComponent
Set Cmp = Md.Parent
End Function
Function CmpTy() As vbext_ComponentType
CmpTy = Cmp.Type
End Function
Function BdyLy() As Sy
Set BdyLy = Sy(SplitCrLf(Lines(BdyLnoCnt)))
End Function
Function ContLin$(Lno&)
Dim L&
L = Lno
Dim O$: O = Md.Lines(L, 1)
While LasChr(O) = "_"
    L = L + 1
    O = RmvLasChr(O) & Md.Lines(L, 1)
Wend
ContLin = O
End Function

Sub Cpy(ToMdNm$, Optional A As VBProject)
Dim ToMd As Md: Set ToMd = ccNew.Pj(A).CrtMd(ToMdNm, CmpTy)
ToMd.AppLy Ly.Sy
End Sub

Function Exp()
Md.Parent.Export SrcFfn
Debug.Print Nm
End Function

Function FunDrs(Optional WithBdyLy As Boolean, Optional WithBdyLines As Boolean) As Drs
Dim ODry As New Dry
    If Not IsEmpty Then
        Dim Dr(), Lin
        Dim MNm$: MNm = Nm
        Dim Lno&, J&
        For Lno = Md.CountOfDeclarationLines + 1 To Md.CountOfLines
            With SrcLin(Lno)
                If .Nm = "" Then GoTo Nxt
                Dr = Array(Lno, .Mdy, .Ty, .Nm, MNm)
                If WithBdyLy Then Push Dr, FunBdyLy(Lno).Sy
                If WithBdyLines Then Push Dr, FunBdyLy(Lno).Lines
                ODry.Push Dr
            End With
Nxt:
        Next
    End If
Dim OFny$()
    OFny = FunDrsFny(WithBdyLy, WithBdyLines)
FunDrs = ccNew.Drs(OFny, ODry)
End Function

Function FunDrsFny(WithBdyLy As Boolean, WithBdyLines As Boolean) As String()
Dim O$(): O = SplitSpc("Lno Mdy Ty FunNm MdNm")
If WithBdyLy Then Push O, "BdyLy"
If WithBdyLines Then Push O, "BdyLines"
FunDrsFny = O
End Function

Function FunLines$(FunNm$)
FunLines = Lines(FunLnoCnt(FunNm))
End Function

Function FunLnoCnt(FunNm$) As LnoCnt
If IsEmpty Then Exit Function
Dim Lno&: Lno = FunLno(FunNm)
If Lno = 0 Then Exit Function
Dim EndLno&: EndLno = FunEndLno(Lno)
Dim Cnt&, O As LnoCnt
Cnt = EndLno - Lno + 1
O.Lno = Lno
O.Cnt = Cnt
FunLnoCnt = O
End Function

Function IsEmpty() As Boolean
IsEmpty = (Md.CountOfLines = 0)
End Function

Function LasLno&()
LasLno = Md.CountOfLines
End Function
Function Lines$(LnoCnt As LnoCnt)
With LnoCnt
    If .Cnt = 0 Then Exit Function
    Lines = Md.Lines(.Lno, .Cnt)
End With
End Function

Function Ly() As Sy
Dim N&: N = Md.CountOfLines
If N = 0 Then Exit Function
Dim O$()
ReDim O(N - 1)
Dim J&
For J = 0 To N - 1
    O(J) = Md.Lines(J + 1, 1)
Next
Set Ly = Sy(O)
End Function

Function Ly_Jn() As Sy
Set Ly_Jn = Sy(JnContinueLin(Ly.Sy))
End Function

Function Nm()
Nm = Md.Parent.Name
End Function

Function Pj() As VBProject
Set Pj = Md.Parent.Collection.Parent
End Function
Sub Rmv()
Cmp.Collection.Remove Cmp
End Sub

Sub RmvBdy()
RmvLnoCnt BdyLnoCnt
End Sub

Sub RmvFun(FunNm$)
Dim M As LnoCnt: M = FunLnoCnt(FunNm)
If M.Cnt = 0 Then
    Debug.Print FmtQQ("Fun[?] in Md[?] not found, cannot Rmv", FunNm, Nm)
Else
    Debug.Print FmtQQ("Fun[?] in Md[?] is removed", FunNm, Nm)
End If
RmvLnoCnt M
End Sub

Sub RmvLnoCnt(LnoCnt As LnoCnt)
If LnoCnt.Cnt = 0 Then Exit Sub
Md.DeleteLines LnoCnt.Lno, LnoCnt.Cnt
End Sub
Function SrcPth$()
SrcPth = Pjx.SrcPth
End Function
Function SrcFfn$()
SrcFfn = SrcPth & SrcFn
End Function

Function SrcFn$()
SrcFn = Cmp.Name & SrcExt
End Function

Sub Srt()
Debug.Print Nm
Dim Lines$: Lines = SrtedBdyLy.Lines
RmvBdy
AppLines Lines
End Sub

Function SrtedBdyLy() As Sy
If IsEmpty Then Exit Function
Dim Drs As Drs: Drs = FunDrs(WithBdyLines:=True)
Dim Ky$()
    Erase Ky
    Dim Ty$, Mdy$, FunNm$, K$, BdyLines$, IdxAy&(), Dr
    IdxAy = FidxAy(Drs.Fny, "Mdy FunNm Ty BdyLines")
    If AyIsEmpty(Drs.Dry) Then Exit Function
    For Each Dr In Drs.Dry
        AyAsg_Idx Dr, IdxAy, Mdy, FunNm, Ty, BdyLines
        Push Ky, FunKey(Mdy, Ty, FunNm)
        If Not IsPfx(LasLin(BdyLines), "End") Then Stop
    Next
Dim I&()
    'AyBrw AySrt(Ky):Stop
    I = AySrtIntoIdxAy(Ky)
Dim O$(), J&
J = 0
Dim IBdyLines%: IBdyLines = AyIdx(Drs.Fny, "BdyLines")
For J = 0 To UB(I)
    Dr = Drs.Dry(I(J))
    BdyLines = Dr(IBdyLines)
    Push O, vbCrLf & BdyLines
Next
Set SrtedBdyLy = Sy(O)
End Function

Function TstFunLines$()
TstFunLines = JnCrLf(TstFunLy)
End Function

Function TstFunLy() As String()
Dim O$(), Ay$()
Ay = TstFunNy
If AyIsEmpty(Ay) Then Exit Function
Push O, "Sub Tst()"
PushAy O, AySrt(TstFunNy)
Push O, "End Sub"
TstFunLy = O
End Function

Function TstFunNy() As String()
If IsEmpty Then Exit Function
Dim M As Drs: M = FunDrs
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
TstFunNy = O
End Function

Function TstFunNy_WithEr() As String()
Dim M As Drs: M = FunDrs
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
TstFunNy_WithEr = O
End Function

Function Ty() As vbext_ComponentType
Ty = Md.Parent.Type
End Function

Sub UpdTstFun()
Dim Md As CodeModule
Set Md = Md
Dim MNm$: MNm = Nm
Select Case CmpTy
Case _
    vbext_ComponentType.vbext_ct_StdModule, _
    vbext_ComponentType.vbext_ct_ClassModule, _
    vbext_ComponentType.vbext_ct_Document, _
    vbext_ComponentType.vbext_ct_MSForm
    Dim NewLines$: NewLines = TstFunLines
    Dim OldLines$: OldLines = FunLines("Tst")
    If OldLines = NewLines Then
        Debug.Print FmtQQ("Fun[Tst] in Md[?] is same: [?] lines", MNm, LinesCnt(OldLines))
        Exit Sub
    End If
    RmvFun "Tst"
    Dim Ly$(): Ly = TstFunLy
    If Sz(Ly) > 0 Then Debug.Print FmtQQ("Fun[Tst] in Md[?] is inserted", MNm)
    AppLy Ly
End Select
End Sub

Function Init(A As CodeModule) As Md
Set Md = A
Set Init = Me
End Function

Private Function FunBdyLy(Lno&) As Sy
Dim EndLno&: EndLno = FunEndLno(Lno)
Dim Cnt&: Cnt = EndLno - Lno + 1
Dim L$
L = Md.Lines(Lno, Cnt)
Set FunBdyLy = Sy(SplitCrLf(L))
End Function

Sub RplLin(Lno&, NewLin$)
Dim OldLin$: OldLin = Md.Lines(Lno, 1)
Stop
End Sub

Private Function FunEndLno&(BegLno&)
Dim Ty$: Ty = SrcLin(BegLno).Ty
Dim Pfx$
    If AyHas(SplitSpc("Get Let Set"), Ty) Then
        Pfx = "End Property"
    Else
        Pfx = "End " & Ty
    End If
Dim O&
    For O = BegLno + 1 To LasLno
        If IsPfx(Md.Lines(O, 1), Pfx) Then FunEndLno = O: Exit Function
    Next
Err.Raise 1, , FmtQQ("Md.FunEndLno: No [?] in module[?] from Lno[?]", Pfx, Nm, BegLno)
End Function

Private Function SrcExt$()
Dim O$
Select Case CmpTy
Case vbext_ct_ClassModule: O = ".cls"
Case vbext_ct_Document: O = ".cls"
Case vbext_ct_StdModule: O = ".bas"
Case Else: Err.Raise 1, , "SrcExt: Unexpected MdCmpTy.  Should be [Class or Module or Document]"
End Select
SrcExt = O
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

Private Sub FunDrs__Tst()
FunDrs(WithBdyLy:=True).Brw
End Sub

Private Sub FunLnoCnt__Tst()
Debug.Print LnoCntToStr(FunLnoCnt("FunLnoCnt"))
End Sub

Private Sub Ly__Tst()
AyBrw Ly
End Sub

Private Sub Srt__Tst()
ccNew.MdByNm("bb_Lib_Acs").Srt
End Sub

Private Sub SrtedBdyLines__Tst()
CurMd.SrtedBdyLy.Ay.Brw
End Sub

Sub Tst()
JnContinueLin__Tst
FunDrs__Tst
FunLnoCnt__Tst
Srt__Tst
SrtedBdyLines__Tst
End Sub



