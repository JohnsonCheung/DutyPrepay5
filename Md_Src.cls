VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Md_Src"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Private A_$()
Public MdNm$
Private GetLetSetAy As Ay
Function Init(Ly$(), MdNm_$) As Md_Src
A_ = Ly
MdNm = MdNm_
Set Init = Me
End Function
Function IsEmpty() As Boolean
IsEmpty = AyIsEmpty(A_)
End Function
Function SrtLy() As String()
Dim mFunDrs As Drs: Set mFunDrs = FunDrs(WithBdyLines:=True, WithRmkLines:=True)
Dim mFunLyAy(): mFunLyAy = FunLyAy(mFunDrs)
Dim mFunKeyAy$(): mFunKeyAy = FunKeyAy(mFunDrs)
Dim mSrtedFunIdxAy&(): mSrtedFunIdxAy = SrtedFunIdxAy(mFunKeyAy)
Dim mDclLy$(): mDclLy = DclLy
Dim O$()
    PushAy O, mDclLy
    Dim J&, Idx&
    For J = 0 To UB(mSrtedFunIdxAy)
        Idx = mSrtedFunIdxAy(J)
        PushAy O, vbCrLf & mFunLyAy(Idx)
    Next
SrtLy = O
End Function
Function FunDrs(Optional WithBdyLy As Boolean, Optional WithBdyLines As Boolean, Optional WithRmkLines As Boolean) As Drs
Dim Fny$(): Fny = FunDrsFny(WithBdyLy, WithBdyLines, WithRmkLines)
Dim Dry As Dry: Set Dry = FunDrsDry(WithBdyLy, WithBdyLines, WithRmkLines)
Set FunDrs = Nw.Drs(Fny, Dry)
End Function
Private Function FunLyAy(FunDrs As Drs) As Variant()
Dim IBdyLines%, IRmkLines%: FunDrs.FldIdxAsg "BdyLines RmkLines", IBdyLines, IRmkLines
Dim Dr, O$(), BdyLines$, RmkLines$, A$
Dim Dry(): Dry = FunDrs.Dry.Dry
For Each Dr In Dry
    BdyLines = Dr(IBdyLines)
    RmkLines = Dr(IRmkLines)
    If RmkLines = "" Then
        A = BdyLines
    Else
        A = RmkLines & vbCrLf & BdyLines
    End If
    Push O, A
    If Not IsPfx(LasLin(A), "End") Then Stop
Next
'FunLyAy = O
End Function
Private Function FunKeyAy(FunDrs As Drs) As String()
Dim Dry
Dim IMdy%
Dim IFunNm%
Dim ITy%

Dim O$(), J%
For J = 0 To FunDrs.URow
'    Push O, FunKey(J)
Next
    'AyBrw AySrt(Ky):Stop
FunKeyAy = O
End Function
Private Function SrtedFunIdxAy(FunKeyAy$()) As Long()
SrtedFunIdxAy = Ay(FunKeyAy).SrtIntoIdxAy
End Function
Function ULin&()
ULin = UB(A_)
End Function
Property Get LinIdx_DclEnd&()
Dim M&: M = LinIdx_FstFun
If M = -1 Then LinIdx_DclEnd = ULin: Exit Property
LinIdx_DclEnd = LinIdx_FstRmkAbove(M) - 1
End Property
Function LinIdx_FstRmkAbove&(FmIdx&)
Dim O&
For O = FmIdx - 1 To 0 Step -1
    With SrcLin(O)
        If Not .IsRmkLin And Not .IsBlank Then LinIdx_FstRmkAbove = O + 1: Exit Function
    End With
Next
LinIdx_FstRmkAbove = 0
End Function
Property Get LinIdx_FstFun&()
Dim O&
For O = 1 To ULin
    If SrcLin(O).IsDfnLin Then LinIdx_FstFun = O: Exit Property
Next
LinIdx_FstFun = O
End Property

Function DclLy() As String()
DclLy = SliceLy(0, LinIdx_DclEnd)
End Function
Function SliceLy(FmIdx&, ToIdx&) As String()
'Slicing the A_
SliceLy = LyX.Slice(FmIdx, ToIdx)
End Function
Private Function IsPfx(S, Pfx) As Boolean
IsPfx = Left(S, Len(Pfx)) = Pfx
End Function

Function SrtLyX() As Ay
Set SrtLyX = Ay(SrtLy)
End Function
Property Get LyX() As Ay
Set LyX = Ay(A_)
End Property

Function SliceLyX(FmIdx&, ToIdx&) As Ay
'Slicing the A_
Set SliceLyX = Ay(SliceLy(FmIdx, ToIdx))
End Function
Function NLin&()
NLin = UB(A_)
End Function

Private Function FunDrsDry(WithBdyLy As Boolean, WithBdyLines As Boolean, WithRmkLines As Boolean) As Dry
Dim O As New Dry
    If IsEmpty Then GoTo N
    Dim Dr(), Lin
    Dim Idx&
    For Idx = 0 To ULin
        With SrcLin(Idx)
            If .Nm = "" Then GoTo Nxt
            Dr = Array(Idx + 1, .Mdy, .Ty, .Nm, MdNm)
            If WithBdyLy Then Push Dr, FunBdyLy(Idx)
            If WithBdyLines Then Push Dr, FunBdyLyX(Idx).Lines
            If WithRmkLines Then Push Dr, FunRmkLyX(Idx).Lines
            ODry.AddDr Dr
        End With
Nxt:
    Next
N:
Set FunDrsDry = O
End Function

Private Function FunDrsFny(WithBdyLy As Boolean, WithBdyLines As Boolean, WithRmkLines As Boolean)
Dim O$(): O = SplitSpc("Lno Mdy Ty FunNm MdNm")
If WithBdyLy Then Push O, "BdyLy"
If WithBdyLines Then Push O, "BdyLines"
If WithRmkLines Then Push O, "RmkLines"
FunDrsFny = O
End Function

Function FunLines$(FunNm$)
FunLines = Lines(FunLIdxCnt(FunNm))
End Function

Private Sub FunLIdxCnt__Tst()
'FunLIdxCnt("FunLIdxCnt").Str
End Sub

Function FunLIdxCnt(FunNm$) As IdxCnt
If IsEmpty Then Exit Function
Dim LIdx&: LIdx = FunLIdx(FunNm)
If LIdx = 0 Then Exit Function
Dim EIdx&: EIdx = FunEIdx(LIdx)
Dim Cnt&, O As New IdxCnt
Cnt = EIdx - LIdx + 1
O.Idx = LIdx
O.Cnt = Cnt
FunLIdxCnt = O
End Function

Private Function FunBdyLy(Idx&) As String()
Dim E&: E = FunEIdx(Idx)
FunBdyLy = SliceLy(Idx, E)
End Function

Private Function FunBdyLyX(Idx&) As Ay
Set FunBdyLyX = Ay(FunBdyLy(Idx))
End Function

Function SrcLin(Idx&) As SrcLin
Set SrcLin = Nw.SrcLin(A_(Idx))
End Function

Private Sub FunRmkLyX__Tst()
ZZ
Dim B&: B = FunLIdx("FunEIdx")
Dim Act$(): Act = FunRmkLyX(B).Ay_
Debug.Assert Sz(Act) = 2
Debug.Assert Act(0) = "'123"
Debug.Assert Act(1) = "'456"
End Sub

Function FunRmkLyX(Idx&) As Ay
Dim B&, E&
B = LinIdx_FstRmkAbove(Idx)
E = Idx - 1
Set FunRmkLyX = SliceLyX(B, E).RmvEmptyEle
End Function

'123


'456
Private Function FunEIdx&(LIdx&)
Dim Ty$: Ty = SrcLin(LIdx).Ty
Dim Pfx$
    If GetLetSetAy.Has(Ty) Then
        Pfx = "End Property"
    Else
        Pfx = "End " & Ty
    End If
Dim O&
    For O = LIdx + 1 To ULin
        If IsPfx(A_(O), Pfx) Then FunEIdx = O: Exit Function
    Next
'Er FmtQQ("MdBdy.FunEIdx: No [?]
End Function

Function Lines$(A As IdxCnt)
With A
Lines = LyX.SliceCntX(.Idx, .Cnt).Lines
End With
End Function
Function TstFunNy() As String()
If IsEmpty Then Exit Function
Dim M As Drs: M = FunDrs
Dim Dr
Dim O$(), Mdy$, Ty$, FunNm$, IFunNm%
IFunNm = M.FldIdx("FunNm")
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
M.AsgFldIdx "Mdy Ty FunNm", IMdy, ITy, IFunNm
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
Function FunLIdx&(FunNm)
If IsEmpty Then Exit Function
Dim O&
Dim J&
For J = 0 To ULin
    With SrcLin(J)
        If .Nm = "" Then GoTo Nxt
        If FunNm = .Nm Then FunLIdx = J: Exit Function
    End With
Nxt:
Next
End Function
Private Function ZZLy() As String()
TstRes.Pth
End Function
Private Sub ZZ()
Init ZZLy, "Testing-MdSrc"
End Sub
Private Sub SrtLyX__Tst()
ZZ
SrtLyX.Brw
End Sub
Sub Tst()
'FunRmkLyX__Tst
'FunDrs__Tst
SrtLyX__Tst
End Sub
Private Sub FunDrs__Tst()
FunDrs(WithBdyLines:=True).Brw
End Sub
Function TstFunLines$()
TstFunLines = JnCrLf(TstFunLy)
End Function

Function TstFunLy() As String()
Dim O$(), N$()
N = TstFunNy
If AyIsEmpty(N) Then Exit Function
Push O, "Sub Tst()"
PushAy O, Ay(TstFunNy).Srt
Push O, "End Sub"
TstFunLy = O
End Function

Private Sub Class_Initialize()
Set GetLetSetAy = Ay(SplitSpc("Get Let Set"))
End Sub
