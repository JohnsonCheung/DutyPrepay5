Attribute VB_Name = "Ide_Src"
Option Compare Database
Option Explicit
Type aa
    A As Integer
End Type

Function SrcAddMthIfNotExist(Src$(), MthNm$, NewMthLy$()) As String()
If SrcHasMth(Src, MthNm) Then
    SrcAddMthIfNotExist = Src
Else
    SrcAddMthIfNotExist = AyAdd(Src, NewMthLy)
End If
End Function

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

Function SrcContLin$(Src$())
SrcContLin = SrcContLinFm(Src, 0)
End Function

Function SrcContLinFm(Src$(), FmIdx&)
If FmIdx = -1 Then Exit Function
Const CSub$ = "SrcContLinFm"
Dim J&, I$
Dim O$, IsCont As Boolean
For J = FmIdx To UB(Src)
    I = Src(J)
    O = O & LTrim(I)
    IsCont = IsSfx(O, " _")
    If IsCont Then O = RmvLasChr(O)
    If Not IsCont Then Exit For
Next
If IsCont Then Er CSub, "each lines {Src} ends with sfx _, which is impossible"
SrcContLinFm = O
End Function

Function SrcDclCnt&(Src$())
Dim I&: I = SrcFstMthIdx(Src)
If I = -1 Then SrcDclCnt = Sz(Src): Exit Function
SrcDclCnt = SrcMthIdxAyFstRmkIdx(Src, I)
End Function

Function SrcDclLines$(Src$())
SrcDclLines = JnCrLf(SrcDclLy(Src))
End Function

Function SrcDclLy(Src$()) As String()
If AyIsEmpty(Src) Then Exit Function
Dim I&
    I = SrcLasDclIdx(Src)
If I = -1 Then Exit Function
SrcDclLy = AyFstUEle(Src, I)
End Function

Function SrcEnmCnt%(Src$())
If AyIsEmpty(Src) Then Exit Function
Dim I, O%
For Each I In Src
    If SrcLinIsEnm(I) Then O = O + 1
Next
SrcEnmCnt = O
End Function

Function SrcEnsTy(Src$(), TyNm$, NewTyLy$()) As String()
Dim OldTyLines$
    OldTyLines = SrcTyLines(Src, TyNm)
Dim NewTyLines$
    NewTyLines = JnCrLf(NewTyLy)
If OldTyLines = NewTyLines Then
    SrcEnsTy = Src
    Exit Function
End If
Dim O$()
    O = SrcRmvTy(Src, TyNm)
    PushAy O, NewTyLy
SrcEnsTy = O
End Function

Function SrcFstMthIdx&(Src$())
Dim J%
For J = 0 To UB(Src)
    If IsMthLin(Src(J)) Then
        SrcFstMthIdx = J
        Exit Function
    End If
Next
SrcFstMthIdx = -1
End Function

Function SrcHasMth(Src$(), MthNm$) As Boolean
SrcHasMth = SrcMthIdx(Src, MthNm) >= 0
End Function

Function SrcLasDclIdx&(Src$())
Dim I&
    I = SrcFstMthIdx(Src)
If I = -1 Then
    SrcLasDclIdx = UB(Src)
Else
    SrcLasDclIdx = I - 1
End If
End Function

Function SrcMthCnt%(Src$())
If AyIsEmpty(Src) Then Exit Function
Dim I, O%
For Each I In Src
    If IsMthLin(I) Then O = O + 1
Next
SrcMthCnt = O
End Function

Function SrcMthDrs(Src$(), Optional WithBdyLy As Boolean, Optional WithBdyLines As Boolean) As Drs
SrcMthDrs.Dry = SrcMthDry(Src, WithBdyLy, WithBdyLines)
SrcMthDrs.Fny = SrcMthDrsFny(WithBdyLy, WithBdyLines)
End Function

Function SrcMthDrsFny(Optional WithBdyLy As Boolean, Optional WithBdyLines As Boolean) As String()
Dim O$(): O = SplitSpc("Lno Mdy Ty MthNm")
If WithBdyLy Then Push O, "BdyLy"
If WithBdyLines Then Push O, "BdyLines"
SrcMthDrsFny = O
End Function

Function SrcMthDry(Src$(), Optional WithBdyLy As Boolean, Optional WithBdyLines As Boolean) As Variant()
Dim MthIdxAy&(): MthIdxAy = SrcMthIdxAyAll(Src)
Dim O()
    Dim Dr()
    Dim J&
    Dim MthIdx&
    Dim BdyLy$()
    For J = 0 To UB(MthIdxAy)
        MthIdx = MthIdxAy(J)
        Dr = SrcLinDr(Src(MthIdx), MthIdx + 1)
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

Function SrcMthIdx&(Src$(), MthNm)
SrcMthIdx = SrcMthIdxFm(Src, MthNm)
End Function

Function SrcMthIdxAy(Src$(), MthNm) As Long()
Dim A&
    A = SrcMthIdx(Src, MthNm)
    If A = -1 Then Exit Function
    
Dim O&()
    Push O, A
    If SrcLinMthTy(Src(A)) <> "Property" Then
        SrcMthIdxAy = O
        Exit Function
    End If
    Dim J%, Fm&, I&
    For J = 1 To 2
        Fm = I + 1
        I = SrcMthIdxFm(Src, MthNm, Fm)
        If I = -1 Then
            SrcMthIdxAy = O
            Exit Function
        End If
        Push O, I
    Next
SrcMthIdxAy = O
End Function

Function SrcMthIdxAyAll(Src$()) As Long()
If AyIsEmpty(Src) Then Exit Function
Dim O&(), J%
For J = 0 To UB(Src)
    If SrcLinMthTy(Src(J)) <> "" Then Push O, J
Next
SrcMthIdxAyAll = O
End Function

Function SrcMthIdxAyFstRmkIdx%(Src$(), MthIdx&)
Dim J%
For J = MthIdx - 1 To 0 Step -1
    If SrcLinIsCd(Src(J)) Then SrcMthIdxAyFstRmkIdx = J + 1: Exit Function
Next
Never
End Function

Function SrcMthIdxFm%(Src$(), MthNm, Optional Fm&)
Dim I%
    For I = Fm To UB(Src)
        If SrcLinMthNm(Src(I)) = MthNm Then
            SrcMthIdxFm% = I
            Exit Function
        End If
    Next
SrcMthIdxFm = -1
End Function

Function SrcMthLin$(Src$(), MthNm$)
SrcMthLin = SrcContLinFm(Src, SrcMthIdx(Src, MthNm))
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
    Push O, ""
Next
SrcMthLy = O
'Function SrcMthLy(Src$(), MthNm$) As String()
'Dim IdxAy&()
'    IdxAy = SrcMthIdxAy(Src, MthNm)
'Dim O$()
'    Dim J%
'    For J = 0 To UB(IdxAy)
'        PushAy O, SrcMthIdxBdyLy(Src, IdxAy(J))
'        Push O, ""
'    Next
'SrcMthLy = O
'End Function
End Function

Function SrcMthNy(Src$(), Optional MthNmLik = "*") As String()
If AyIsEmpty(Src) Then Exit Function
Dim O$(), L, M$
For Each L In Src
    M = SrcLinBrk(L).MthNm
    If M Like MthNmLik Then
        PushNonEmpty O, M
    End If
Next
SrcMthNy = O
End Function

Function SrcPrvMthNy(Src$()) As String()
If AyIsEmpty(Src) Then Exit Function
Dim O$(), L
For Each L In Src
    With SrcLinBrk(L)
        If .Mdy = "Private" Then
            PushNonEmpty O, .MthNm
        End If
    End With
Next
SrcPrvMthNy = O
End Function

Function SrcRmvMth(Src$(), MthNm$) As String()
Dim FmToAy() As FmTo
    FmToAy = SrcMthFmToAy(Src, MthNm)
Dim O$()
    O = Src
    Dim J%
    For J = FmToUB(FmToAy) To 0 Step -1
        O = AyRmvFmTo(O, FmToAy(J))
    Next
SrcRmvMth = O
End Function

Function SrcRmvTy(Src$(), TyNm$) As String()
SrcRmvTy = AyRmvFmTo(Src, SrcTyFmTo(Src, TyNm))
End Function

Function SrcRplMth(Src$(), MthNm$, NewMthLy$()) As String()
Dim OldMthLines$
    OldMthLines = SrcMthLines(Src, MthNm)
Dim NewMthLines$
    NewMthLines = JnCrLf(NewMthLy)
If OldMthLines = NewMthLines Then
    SrcRplMth = Src
    Exit Function
End If
Dim O$()
    O = SrcRmvMth(Src, MthNm)
    PushAy O, NewMthLy
SrcRplMth = O

End Function

Function SrcSrtedBdyLines$(Src$())
If AyIsEmpty(Src) Then Exit Function
Dim Drs As Drs
    Drs = SrcMthDrs(Src, WithBdyLines:=True)

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
SrcSrtedBdyLines = JnCrLf(O)
End Function

Function SrcSrtLines$(Src$())
SrcSrtLines = SrcDclLines(Src) & SrcSrtedBdyLines(Src)
End Function

Function SrcSrtLy(Src$()) As String()
SrcSrtLy = SplitCrLf(SrcSrtLines(Src))
End Function

Function SrcTyCnt%(Src$())
If AyIsEmpty(Src) Then Exit Function
Dim I, O%
For Each I In Src
    If SrcLinIsTy(I) Then O = O + 1
Next
SrcTyCnt = O
End Function

Function SrcTyEndIdx(Src$(), FmI&)
Dim O&
For O = FmI + 1 To UB(Src)
    If IsPfx(Src(O), "End Type") Then SrcTyEndIdx = O: Exit Function
Next
SrcTyEndIdx = -1
End Function

Function SrcTyFmIdx&(Src$(), TyNm$)
Dim J%, L$
For J = 0 To UB(Src)
    L = RmvPfx(Src(J), "Private ")
    If IsPfx(L, "Type ") Then SrcTyFmIdx = J: Exit Function
Next
SrcTyFmIdx = -1
End Function

Function SrcTyFmTo(Src$(), TyNm$) As FmTo
Dim FmI&: FmI = SrcTyFmIdx(Src, TyNm)
Dim ToI&: ToI = SrcTyEndIdx(Src, FmI)
SrcTyFmTo = FmTo(FmI, ToI)
End Function

Function SrcTyLines$(Src$(), TyNm$)
SrcTyLines = JnCrLf(SrcTyLy(Src, TyNm))
End Function

Function SrcTyLy(Src$(), TyNm$) As String()
Dim DclLy$()
    DclLy = SrcDclLy(Src)
SrcTyLy = AySelFmTo(Src, SrcTyFmTo(DclLy, TyNm))
End Function

Private Function SrcMthEndIdx&(Src$(), MthIdx&)
Const CSub = "SrcMthEndIdx"
Dim MthLin$
    MthLin = Src(MthIdx)
    
Dim Pfx$
    Pfx = MthLinEndLinPfx(MthLin)
Dim O&
    For O = MthIdx + 1 To UB(Src)
        If IsPfx(Src(O), Pfx) Then SrcMthEndIdx = O: Exit Function
    Next
Er CSub, "{Src}-{MthIdx} is {MthLin} which does have {FunEndLinPfx} in lines after [MthIdx]", Src, MthIdx, MthLin, Pfx
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

Private Sub SrcContLin__Tst()
Dim O$(3)
O(0) = "A _"
O(1) = "  B _"
O(2) = "C"
O(3) = "D"
Dim Act$: Act = SrcContLin(O)
Debug.Assert Act = "A B C"
End Sub

Private Sub SrcDclCnt__Tst()
Dim Src$(): Src = MdSrc(Md("LnkT"))
Debug.Assert SrcDclCnt(Src) = 15
End Sub

Private Sub SrcFstMthIdx__Tst()
Dim Src$(): Src = MdSrc
Debug.Assert SrcFstMthIdx(Src) = 19
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

Private Sub SrcTyLines__Tst()
Debug.Print SrcTyLines(MdSrc, "AA")
End Sub
