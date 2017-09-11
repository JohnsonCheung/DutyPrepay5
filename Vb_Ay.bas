Attribute VB_Name = "Vb_Ay"
Option Explicit
Option Compare Database

Type FmTo
    FmIdx As Long
    ToIdx As Long
End Type

Sub AssertChk(Chk$())
If AyIsEmpty(Chk) Then Exit Sub
AyBrw Chk
Err.Raise 1, , "Error checked!"
End Sub

Function AyAdd(Ay, ParamArray AyAp())
Dim Av(): Av = AyAp
Dim A
Dim O: O = Ay
For Each A In Av
    PushAy O, A
Next
AyAdd = O
End Function

Function AyAddOneAy(Ay1, Ay2)
Dim O: O = Ay1
PushAy O, Ay2
AyAddOneAy = O
End Function

Function AyAddPfx(Ay, Pfx) As String()
Dim O$(), U&, J&
U = UB(Ay)
If U = -1 Then Exit Function
ReDim Preserve O(U)
For J = 0 To U
    O(J) = Pfx & Ay(J)
Next
AyAddPfx = O
End Function

Function AyAddPfxSfx(Ay, Pfx, Sfx) As String()
Dim O$(), U&, J&
U = UB(Ay)
If U = -1 Then Exit Function
ReDim Preserve O(U)
For J = 0 To U
    O(J) = Pfx & Ay(J) & Sfx
Next
AyAddPfxSfx = O
End Function

Function AyAddSfx(Ay, Sfx) As String()
Dim O$(), U&, J&
U = UB(Ay)
If U = -1 Then Exit Function
ReDim Preserve O(U)
For J = 0 To U
    O(J) = Ay(J) & Sfx
Next
AyAddSfx = O
End Function

Function AyAlignL(Ay) As String()
If AyIsEmpty(Ay) Then Exit Function
Dim W%: W = AyWdt(Ay)
Dim O$(), I
For Each I In Ay
    Push O, AlignL(I, W)
Next
AyAlignL = O
End Function

Sub AyAsg(Ay, ParamArray OAp())
Dim Av(): Av = OAp
Dim J%
For J = 0 To UB(Av)
    If Not IsMissing(OAp(J)) Then
        OAp(J) = Ay(J)
    End If
Next
End Sub

Sub AyAsg_Idx(Ay, IdxAy&(), ParamArray OAp())
Dim J%
For J = 0 To UB(IdxAy)
    OAp(J) = Ay(IdxAy(J))
Next
End Sub

Function AyAsgInto(Ay, OIntoAy)
If AyIsEmpty(Ay) Then
    Erase OIntoAy
    AyAsgInto = OIntoAy
    Exit Function
End If
Dim U&
    U = UB(Ay)
ReDim OIntoAy(U)
Dim I, J&
For Each I In Ay
    Asg I, OIntoAy(J)
    J = J + 1
Next
AyAsgInto = OIntoAy
End Function

Sub AyAssertEq(Ay1, Ay2, Optional Ay1Nm$ = "Exp", Optional Ay2Nm$ = "Act")
AssertChk ChkEqAy(Ay1, Ay2, Ay1Nm, Ay2Nm)
End Sub

Function AyBrw(Ay)
Dim T$
T = TmpFt
AyWrt Ay, T
FtBrw T
End Function

Function AyDist(Ay)
Dim O: O = Ay: Erase O
Dim I
For Each I In Ay
    PushNoDup O, I
Next
AyDist = O
End Function

Sub AyDmp(Ay)
If AyIsEmpty(Ay) Then Exit Sub
Dim I
Dim J&
For Each I In Ay
    Debug.Print J; ": ["; I; "]"
    J = J + 1
Next
End Sub

Function AyDupAy(Ay)
Dim O: O = Ay: Erase O
Dim GpDry(): GpDry = AyGpDry(Ay)
If AyIsEmpty(GpDry) Then AyDupAy = O: Exit Function
Dim Dr
For Each Dr In GpDry
    If Dr(1) > 1 Then Push O, Dr(0)
Next
AyDupAy = O
End Function

Function AyFilter(Ay, FilterFunNm$, ParamArray Ap())
Dim O: O = Ay: Erase O
Dim I
Dim Av()
    Av = Ap
    AyIns Av
For Each I In Ay
    Asg I, Av(0)
    If RunAv(FilterFunNm, Av) Then
        Push O, I
    End If
Next
AyFilter = O
End Function

Function AyFm(Ay, FmIdx&)
Dim O: O = Ay: Erase O
If 0 <= FmIdx And FmIdx <= UB(Ay) Then
    Dim J&
    For J = FmIdx To UB(Ay)
        Push O, Ay(J)
    Next
End If
AyFm = O
End Function

Function AyFmTo(Ay, FmTo As FmTo)
Dim O: O = Ay: Erase O
Dim J&
For J = FmTo.FmIdx To FmTo.ToIdx
    Push O, Ay(J)
Next
AyFmTo = O
End Function

Function AyGpDry(Ay) As Variant()
If AyIsEmpty(Ay) Then Exit Function
Dim O(), I
For Each I In Ay
    GpDryUpd O, I
Next
AyGpDry = O
End Function

Function AyHas(Ay, Itm) As Boolean
If AyIsEmpty(Ay) Then Exit Function
Dim I
For Each I In Ay
    If I = Itm Then AyHas = True: Exit Function
Next
End Function

Function AyIdx&(Ay, Itm)
Dim J&
For J = 0 To UB(Ay)
    If Ay(J) = Itm Then AyIdx = J: Exit Function
Next
AyIdx = -1
End Function

Function AyIdxAy(Ay, SubAy) As Long()
If AyIsEmpty(SubAy) Then Exit Function
Dim O&()
Dim U&: U = UB(SubAy)
ReDim O(U)
Dim J&
For J = 0 To U
    O(J) = AyIdx(Ay, SubAy(J))
Next
AyIdxAy = O
End Function

Sub AyIns(OAy, Optional Ele, Optional At&)
Dim N&: N = Sz(OAy)
If 0 > At Or At > N Then Err.Raise 1, , FmtQQ("At[?] is outside OAy-UB[?]", At, UB(OAy))
ReDim Preserve OAy(N)
Dim J&
For J = N To At + 1 Step -1
    OAy(J) = OAy(J - 1)
Next
OAy(At) = Ele
End Sub

Function AyIsEmpty(Ay) As Boolean
AyIsEmpty = (Sz(Ay) = 0)
End Function

Function AyIsEq(Ay1, Ay2) As Boolean
Dim U&: U = UB(Ay1): If U <> UB(Ay2) Then Exit Function
Dim J&
For J = 0 To U
    If Ay1(J) <> Ay2(J) Then Exit Function
Next
AyIsEq = True
End Function

Function AyLasEle(Ay)
AyLasEle = Ay(UB(Ay))
End Function

Function AyMap(Ay, FunNm$, ParamArray Ap()) As Variant()
If AyIsEmpty(Ay) Then Exit Function
Dim Av(): Av = Ap
AyIns Av
Dim I, J&
Dim O()
Dim U&: U = UB(Ay)
    ReDim O(U)
For Each I In Ay
    Asg I, Av(0)
    Asg RunAv(FunNm, Av), O(J)
    J = J + 1
Next
AyMap = O
End Function

Function AyMapIntoSy(Ay, FunNm$, ParamArray Ap()) As String()
If AyIsEmpty(Ay) Then Exit Function
Dim Av(): Av = Ap
If AyIsEmpty(Av) Then
    AyMapIntoSy = AyMapIntoSy_NoPrm(Ay, FunNm)
    Exit Function
End If
Dim I, J&
Dim O$()
    ReDim O(UB(Ay))
    AyIns Av
    For Each I In Ay
        Asg I, Av(0)
        Asg RunAv(FunNm, Av), O(J)
        J = J + 1
    Next
AyMapIntoSy = O
End Function

Function AyMapIntoSy_NoPrm(Ay, FunNm$) As String()
If AyIsEmpty(Ay) Then Exit Function
Dim I, J&
Dim O$()
    ReDim O(UB(Ay))
    For Each I In Ay
        Asg Run(FunNm, I), O(J)
        J = J + 1
    Next
AyMapIntoSy_NoPrm = O
End Function

Function AyMinus(Ay, ParamArray AyAp())
Dim O: O = Ay
Dim Av(): Av = AyAp
Dim Ay1, V
For Each Ay1 In Av
    If AyIsEmpty(O) Then AyMinus = O: Exit Function
    O = AyMinusOneAy(O, Ay1)
Next
AyMinus = O
End Function

Function AyMinusOneAy(Ay1, Ay2)
If AyIsEmpty(Ay1) Then AyMinusOneAy = Ay1: Exit Function
Dim O: O = Ay1: Erase O
Dim mAy2: mAy2 = Ay2
Dim V
For Each V In Ay1
    If AyHas(mAy2, V) Then
        AyRmvEle mAy2, V
    Else
        Push O, V
    End If
Next
AyMinusOneAy = O
End Function

Function AyQuote(Ay, QuoteStr$) As String()
If AyIsEmpty(Ay) Then Exit Function
Dim U&: U = UB(Ay)
Dim O$()
    ReDim O(U)
    Dim J&
    Dim Q1$, Q2$
    With BrkQuote(QuoteStr)
        Q1 = .S1
        Q2 = .S2
    End With
    For J = 0 To U
        O(J) = Q1 & Ay(J) & Q2
    Next
AyQuote = O
End Function

Sub AyRmvEle(Ay, Optional Ele, Optional At& = -1)
Dim Idx&
    If IsMissing(Ele) Then
        Idx = At
    Else
        Idx = AyIdx(Ay, Ele)
    End If
AyRmvEleAtCnt Ay, Idx, 1
End Sub

Sub AyRmvEleAtCnt(Ay, At&, Optional Cnt& = 1)
If Cnt <= 0 Then Exit Sub
If At < 0 Then Exit Sub
Dim U&: U = UB(Ay)
If At > U Then Exit Sub
If U = 0 Then Exit Sub
Dim J&
For J = At To U - Cnt
    Ay(J) = Ay(J + Cnt)
Next
ReDim Preserve Ay(U - Cnt)
End Sub

Sub AyRmvEmptyEleAtEnd(OAy)
Dim LasU&, U&
For LasU = UB(OAy) To 0 Step -1
    If Not IsEmpty(OAy(LasU)) Then
        Exit For
    End If
Next
If LasU = -1 Then
    Erase OAy
Else
    ReDim Preserve OAy(LasU)
End If
End Sub

Sub AyRmvFstEle(Ay)
AyRmvEle Ay, At:=0
End Sub

Sub AyRmvLasEle(Ay)
AyRmvEle Ay, At:=UB(Ay)
End Sub

Function AySelByIdxAy(Ay, IdxAy%())
Dim O: O = Ay: Erase O
Dim J&
For J = 0 To UB(IdxAy)
    Push O, Ay(IdxAy(J))
Next
AySelByIdxAy = O
End Function

Function AyShift(Ay)
AyShift = Ay(0)
AyRmvFstEle Ay
End Function

Function AySrt(Ay, Optional Des As Boolean)
If AyIsEmpty(Ay) Then AySrt = Ay: Exit Function
Dim Idx&, V, J&
Dim O: O = Ay: Erase O
Push O, Ay(0)
For J = 1 To UB(Ay)
    AyIns O, Ay(J), AySrt__Idx(O, Ay(J), Des)
Next
AySrt = O
End Function

Function AySrtIntoIdxAy(Ay, Optional Des As Boolean) As Long()
If AyIsEmpty(Ay) Then Exit Function
Dim Idx&, V, J&
Dim O&():
Push O, 0
For J = 1 To UB(Ay)
    AyIns O, J, AySrtIntoIdxAy_Idx(O, Ay, Ay(J), Des)
Next
AySrtIntoIdxAy = O
End Function

Function AyStrAy(Ay) As String()
If AyIsEmpty(Ay) Then Exit Function
Dim U&: U = UB(Ay)
Dim O$()
    Dim J&
    ReDim O(U)
    For J = 0 To U
        O(J) = Ay(J)
    Next
AyStrAy = O
End Function

Function AySy(Ay) As String()
If AyIsEmpty(Ay) Then Exit Function
If IsStrAy(Ay) Then AySy = Ay: Exit Function
Dim U&, O$(), J&, I
J = 0
U = UB(Ay)
ReDim O(U)
For Each I In Ay
    O(J) = I
    J = J + 1
Next
AySy = O
End Function

Sub AyTrim(Ay)
Dim J&
For J = 0 To UB(Ay)
    Ay(J) = Trim(Ay(J))
Next
End Sub

Function AyWdt%(Ay)
Dim O%, I
For Each I In Ay
    O = Max(O, Len(I))
Next
AyWdt = O
End Function

Sub AyWrt(Ay, Ft)
StrWrt JnCrLf(Ay), Ft
End Sub

Function ChkEqAy(Ay1, Ay2, Optional Ay1Nm$ = "Exp", Optional Ay2Nm$ = "Act") As String()
Dim U&: U = UB(Ay1)
Dim O$()
    If U <> UB(Ay2) Then Push O, FmtQQ("Array [?] and [?] has different Sz: [?] [?]", Ay1Nm, Ay2Nm, Sz(Ay1), Sz(Ay2)): GoTo X
If AyIsEmpty(Ay1) Then Exit Function
Dim O1$()
    Dim A2: A2 = Ay2
    Dim J&, ReachLimit As Boolean
    Dim Cnt%
    For J = 0 To U
        If Ay1(J) <> Ay2(J) Then
            Push O1, FmtQQ("[?]-th Ele is diff: ?[?]<>?[?]", Ay1Nm, Ay2Nm, Ay1(J), Ay2(J))
            Cnt = Cnt + 1
        End If
        If Cnt > 10 Then
            ReachLimit = True
            Exit For
        End If
    Next
If IsEmpty(O1) Then Exit Function
Dim O2$()
    Push O2, FmtQQ("Array [?] and [?] both having size[?] have differnt element(s):", Ay1Nm, Ay2Nm, Sz(Ay1))
    If ReachLimit Then
        Push O2, FmtQQ("At least [?] differences:", Sz(O1))
    End If
PushAy O, O2
PushAy O, O1
X:
Push O, FmtQQ("Ay-[?]:", Ay1Nm)
PushAy O, AyQuote(Ay1, "[]")
Push O, FmtQQ("Ay-[?]:", Ay2Nm)
PushAy O, AyQuote(Ay2, "[]")
ChkEqAy = O
End Function

Function DblAyQuote(Ay) As String()
Dim O$(), U&, J&
U = UB(Ay)
ReDim Preserve O(U)
For J = 0 To U
    O(J) = """" & Ay(J) & """"
Next
DblAyQuote = O
End Function

Function EmptyAy() As Variant()
End Function

Function EmptyFmTo() As FmTo
EmptyFmTo.FmIdx = -1
EmptyFmTo.ToIdx = -1
End Function

Function EmptyIntAy() As Integer()
End Function

Function EmptySy() As String()
End Function

Function Mul2&(A)
Mul2 = A * 2
End Function

Function Pop(Ay)
Pop = AyLasEle(Ay)
RmvLasNEle Ay
End Function

Sub Push(O, P)
Dim N&: N = Sz(O)
ReDim Preserve O(N)
If IsObject(P) Then
    Set O(N) = P
Else
    O(N) = P
End If
End Sub

Sub PushAy(OAy, Ay)
If AyIsEmpty(Ay) Then Exit Sub
Dim I
For Each I In Ay
    Push OAy, I
Next
End Sub

Sub PushNoDup(Ay, I)
If Not AyHas(Ay, I) Then Push Ay, I
End Sub

Sub PushNonEmpty(Ay, I)
If IsEmpty(I) Then Exit Sub
Push Ay, I
End Sub

Sub PushObj(O, P)
Dim N&: N = Sz(O)
ReDim Preserve O(N)
Set O(N) = P
End Sub

Sub RmvLasNEle(Ay, Optional NEle% = 1)
ReDim Preserve Ay(UB(Ay) - NEle)
End Sub

Function RunAv(FunNm$, Av())
Dim O
Select Case Sz(Av)
Case 0: O = Run(FunNm)
Case 1: O = Run(FunNm, Av(0))
Case 2: O = Run(FunNm, Av(0), Av(1))
Case 3: O = Run(FunNm, Av(0), Av(1), Av(2))
Case 4: O = Run(FunNm, Av(0), Av(1), Av(2), Av(3))
Case 5: O = Run(FunNm, Av(0), Av(1), Av(2), Av(3), Av(4))
Case 6: O = Run(FunNm, Av(0), Av(1), Av(2), Av(3), Av(4), Av(5))
Case 7: O = Run(FunNm, Av(0), Av(1), Av(2), Av(3), Av(4), Av(5), Av(6))
Case 8: O = Run(FunNm, Av(0), Av(1), Av(2), Av(3), Av(4), Av(5), Av(6), Av(7))
Case 9: O = Run(FunNm, Av(0), Av(1), Av(2), Av(3), Av(4), Av(5), Av(6), Av(7), Av(8))
Case Else: Stop
End Select
RunAv = O
End Function

Function StrWrt(S, Ft)
Dim F%: F = FreeFile(1)
Open Ft For Output As #F
Print #F, S
Close #F
'Dim T As TextStream
'Set T = Fso.OpenTextFile(Ft, ForWriting, True)
'T.Write S
'T.Close
End Function

Function Sy(ParamArray Ap()) As String()
Dim Av(): Av = Ap
Sy = AySy(Av)
End Function

Function Sz&(Ay)
On Error Resume Next
Sz = UBound(Ay) + 1
End Function

Function UB&(Ay)
UB = Sz(Ay) - 1
End Function

Private Function AySrt__Idx&(Ay, V, Des As Boolean)
Dim I, O&
If Des Then
    For Each I In Ay
        If V > I Then AySrt__Idx = O: Exit Function
        O = O + 1
    Next
    AySrt__Idx = O
    Exit Function
End If
For Each I In Ay
    If V < I Then AySrt__Idx = O: Exit Function
    O = O + 1
Next
AySrt__Idx = O
End Function

Private Function AySrtIntoIdxAy_Idx&(Idx&(), Ay, V, Des As Boolean)
Dim I, O&
If Des Then
    For Each I In Idx
        If V > Ay(I) Then AySrtIntoIdxAy_Idx& = O: Exit Function
        O = O + 1
    Next
    AySrtIntoIdxAy_Idx& = O
    Exit Function
End If
For Each I In Idx
    If V < Ay(I) Then AySrtIntoIdxAy_Idx& = O: Exit Function
    O = O + 1
Next
AySrtIntoIdxAy_Idx& = O
End Function

Private Sub GpDryUpd(OGpDry(), Itm)
Dim J&
For J = 0 To UB(OGpDry)
    If OGpDry(J)(0) = Itm Then
        OGpDry(J)(1) = OGpDry(J)(1) + 1
        Exit Sub
    End If
Next
Push OGpDry, Array(Itm, 1)
End Sub

Private Sub TstResPthBrw()
PthBrw TstResPth
End Sub

Sub AyAddOneAy__Tst()
Dim Act(), Exp(), Ay1(), Ay2()
Ay1 = Array(1, 2, 2, 2, 4, 5)
Ay2 = Array(2, 2)
Act = AyAddOneAy(Ay1, Ay2)
Exp = Array(1, 2, 2, 2, 4, 5, 2, 2)
AyAssertEq Exp, Act
AyAssertEq Ay1, Array(1, 2, 2, 2, 4, 5)
AyAssertEq Ay2, Array(2, 2)
End Sub

Private Sub AyGpDry__Tst()
Dim Ay$(): Ay = SplitSpc("a a a b c b")
Dim Act(): Act = AyGpDry(Ay)
Dim Exp(): Exp = Array(Array("a", 3), Array("b", 2), Array("c", 1))
DryAssertEq Act, Exp
End Sub

Private Sub AyMap__Tst()
Dim Act: Act = AyMap(Array(1, 2, 3, 4), "Mul2")
Debug.Assert Sz(Act) = 4
Debug.Assert Act(0) = 2
Debug.Assert Act(1) = 4
Debug.Assert Act(2) = 6
Debug.Assert Act(3) = 8
End Sub

Sub AyMinusOneAy__Tst()
Dim Act(), Exp()
Dim Ay1(), Ay2()
Ay1 = Array(1, 2, 2, 2, 4, 5)
Ay2 = Array(2, 2)
Act = AyMinusOneAy(Ay1, Ay2)
Exp = Array(1, 2, 4, 5)
AyAssertEq Exp, Act
'
Act = AyMinus(Array(1, 2, 2, 2, 4, 5), Array(2, 2), Array(5))
Exp = Array(1, 2, 4)
AyAssertEq Exp, Act
End Sub

Private Sub AyRmvEleAtCnt__Tst()
Dim Ay(): Ay = Array(1, 2, 3, 4, 5)
AyRmvEleAtCnt Ay, 1, 2
AyAssertEq Array(1, 4, 5), Ay
End Sub

Private Sub AyRmvEmptyEleAtEnd__Tst()
Dim Ay: Ay = Array(Empty, Empty, Empty, 1, Empty, Empty)
AyRmvEmptyEleAtEnd Ay: Debug.Assert Sz(Ay) = 4
End Sub

Private Sub AySrt__Tst()
Dim Exp, Act
Dim Ay
Ay = Array(1, 2, 3, 4, 5): Exp = Ay:                   Act = AySrt(Ay):       AyAssertEq Exp, Act
Ay = Array(1, 2, 3, 4, 5): Exp = Array(5, 4, 3, 2, 1): Act = AySrt(Ay, True): AyAssertEq Exp, Act
Ay = Array(":", "~", "P"): Exp = Array(":", "P", "~"): Act = AySrt(Ay):       AyAssertEq Exp, Act
'-----------------
Erase Ay
Push Ay, ":PjUpdTstFun:Sub"
Push Ay, ":SrcLinBrk:Function"
Push Ay, "~~:Tst:Sub"
Push Ay, ":PjTstFunNy_WithEr:Function"
Push Ay, "~Private:JnContinueLin__Tst:Sub"
Push Ay, "Private:IsPfx:Function"
Push Ay, "Private:MdFunDrs_FunBdyLy:Function"
Push Ay, "Private:SrcFunIdxEndIdx:Function"
Erase Exp
Push Exp, ":PjTstFunNy_WithEr:Function"
Push Exp, ":PjUpdTstFun:Sub"
Push Exp, ":SrcLinBrk:Function"
Push Exp, "Private:IsPfx:Function"
Push Exp, "Private:MdFunDrs_FunBdyLy:Function"
Push Exp, "Private:SrcFunIdxEndIdx:Function"
Push Exp, "~Private:JnContinueLin__Tst:Sub"
Push Exp, "~~:Tst:Sub"
Act = AySrt(Ay):       AyAssertEq Exp, Act
'---------------------
Ay = FtLy(TstResPth & "AySrt_Ft1.txt")
Exp = FtLy(TstResPth & "AySrt_Ft1_Exp.txt")
Act = AySrt(Ay):       AyAssertEq Exp, Act

End Sub

Private Sub AySrtIntoIdxAy__Tst()
Dim Ay: Ay = Array("A", "B", "C", "D", "E")
AyAssertEq Array(0, 1, 2, 3, 4), AySrtIntoIdxAy(Ay)
AyAssertEq Array(4, 3, 2, 1, 0), AySrtIntoIdxAy(Ay, True)
End Sub

Private Sub ChkEqAy__Tst()
AyDmp ChkEqAy(Array(1, 2, 3, 3, 4), Array(1, 2, 3, 4, 4))
End Sub

Sub Tst()
AyAddOneAy__Tst
AyMap__Tst
AyMinusOneAy__Tst
AyRmvEleAtCnt__Tst
AyRmvEmptyEleAtEnd__Tst
AySrt__Tst
AySrtIntoIdxAy__Tst
ChkEqAy__Tst
End Sub
