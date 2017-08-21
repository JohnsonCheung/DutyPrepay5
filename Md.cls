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
EnmBdyLyX("AA").Dmp
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
Function EnmBdyLyX(EnmNm$) As Ay
Set EnmBdyLyX = Ay(EnmBdyLy(EnmNm))
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
Sub BrwSrcPth()
PjX.SrcPthX.Brw
End Sub
Function PjX() As Pj
Set PjX = Nw.Pj(Pj)
End Function

Sub AppLines(Lines$)
If Lines = "" Then Exit Sub
Md.InsertLines LasLno + 1, Lines
End Sub

Sub AppLy(Ly$())
AppLines JnCrLf(Ly)
End Sub
Function BdyFstLno&()
BdyFstLno = Md.CountOfDeclarationLines + 1
End Function
Function SrcIdxCnt() As IdxCnt
Dim O As IdxCnt
O.Idx = 1
O.Cnt = Md.CountOfLines
Set SrcIdxCnt = O
End Function
Function Src() As Md_Src
Dim O As New Md_Src
Set Src = O.Init(Ly, Nm)
End Function
Function Cmp() As VBComponent
Set Cmp = Md.Parent
End Function
Function CmpTy() As vbext_ComponentType
CmpTy = Cmp.Type
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
Dim ToMd As Md: Set ToMd = Nw.Pj(A).CrtMd(ToMdNm, CmpTy)
ToMd.AppLy Ly
End Sub

Function Exp()
Md.Parent.Export SrcFfn
Debug.Print Nm
End Function

Function IsEmpty() As Boolean
IsEmpty = (Md.CountOfLines = 0)
End Function

Function LasLno&()
LasLno = Md.CountOfLines
End Function
Function Lines$(IdxCnt As IdxCnt)
With IdxCnt
    If .Cnt = 0 Then Exit Function
    Lines = Md.Lines(.Idx + 1, .Cnt)
End With
End Function

Function Ly() As String()
Dim N&: N = Md.CountOfLines
If N = 0 Then Exit Function
Dim O$()
ReDim O(N - 1)
Dim J&
For J = 0 To N - 1
    O(J) = Md.Lines(J + 1, 1)
Next
Ly = O
End Function

Function LyNoContLin() As String()
Dim O$(): O = Ly
Dim J&
For J = UB(O) - 1 To 0 Step -1
    If LasChr(O(J)) = "_" Then
        O(J) = RmvLasChr(O(J)) & O(J + 1)
        O(J + 1) = ""
    End If
Next
LyNoContLin = O
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

Sub RmvSrc()
RmvIdxCnt SrcIdxCnt
End Sub

Sub RmvFun(FunNm$)
Dim M As IdxCnt: M = Src.FunLIdxCnt(FunNm)
If M.Cnt = 0 Then
    Debug.Print FmtQQ("Fun[?] in Md[?] not found, cannot Rmv", FunNm, Nm)
Else
    Debug.Print FmtQQ("Fun[?] in Md[?] is removed", FunNm, Nm)
End If
RmvIdxCnt M
End Sub

Sub RmvIdxCnt(IdxCnt As IdxCnt)
If IdxCnt.Cnt = 0 Then Exit Sub
Md.DeleteLines IdxCnt.Idx + 1, IdxCnt.Cnt
End Sub
Function SrcPth$()
SrcPth = PjX.SrcPth
End Function
Function SrcFfn$()
SrcFfn = SrcPth & SrcFn
End Function

Function SrcFn$()
SrcFn = Cmp.Name & SrcExt
End Function

Sub Srt()
Debug.Print Nm
Dim Lines$: Lines = Src.SrtLyX.Lines
RmvSrc
AppLines Lines
End Sub
Private Function IsPfx(S, Pfx) As Boolean
IsPfx = (Left(S, Len(Pfx)) = Pfx)
End Function
Sub MdCrt(Optional MdNm$, Optional Ty As vbext_ComponentType = vbext_ct_StdModule, Optional A As VBProject)
Dim O As VBComponent: Set O = DftPj(A).VBComponents.Add(Ty)
O.CodeModule.DeleteLines 1, 2
If MdNm <> "" Then O.Name = MdNm
End Sub
Function DftPjNm$(Optional PjNm$)
If PjNm = "" Then
    DftPjNm = DftPj.Name
Else
    DftPjNm = PjNm
End If
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
    With Src
    Dim NewLines$: NewLines = .TstFunLines
    Dim OldLines$: OldLines = .FunLines("Tst")
    End With
    If OldLines = NewLines Then
        Debug.Print FmtQQ("Fun[Tst] in Md[?] is same: [?] lines", MNm, LinesCnt(OldLines))
        Exit Sub
    End If
    RmvFun "Tst"
    Dim Ly$(): Ly = Src.TstFunLy
    If Sz(Ly) > 0 Then Debug.Print FmtQQ("Fun[Tst] in Md[?] is inserted", MNm)
    AppLy Ly
End Select
End Sub

Function Init(Optional Md_) As Md
Set Md = DftMd(Md_)
Set Init = Me
End Function


Sub RplLin(Lno&, NewLin$)
Dim OldLin$: OldLin = Md.Lines(Lno, 1)
Stop
End Sub

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
Private Function ZZ() As Md
End Function
Private Sub LyX__Tst()
ZZ.LyX.Brw
End Sub
Function LyX() As Ay
Set LyX = Ay(Ly)
End Function
Private Sub Srt__Tst()
Nw.MdByNm("bb_Lib_Acs").Srt
End Sub


Private Sub LyNoContLin__Tst _
()
Dim O$(3)
O(0) = "A _"
O(1) = "B _"
O(2) = "C"
O(3) = _
  "D"
Dim Act$(): Act = ZZ.LyNoContLin
Debug.Assert UB(Act) = 3
Debug.Assert Act(0) = "A B C"
Debug.Assert Act(1) = ""
Debug.Assert Act(2) = ""
Debug.Assert Act(3) = "D"
End Sub
