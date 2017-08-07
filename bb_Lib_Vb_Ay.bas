Attribute VB_Name = "bb_Lib_Vb_Ay"
Option Compare Database
Option Explicit
Function AyIsEmpty(Ay) As Boolean
AyIsEmpty = (Sz(Ay) = 0)
End Function
Sub AyAsg_Idx(Ay, IdxAy&(), ParamArray OAp())
Dim J%
For J = 0 To UB(IdxAy)
    OAp(J) = Ay(IdxAy(J))
Next
End Sub
Sub AyAsg(Ay, ParamArray OAp())
Dim Av(): Av = OAp
Dim J%
For J = 0 To UB(Av)
    If Not IsMissing(OAp(J)) Then
        OAp(J) = Ay(J)
    End If
Next
End Sub

Sub AyIns(OAy, Ele, Optional At&)
Dim N&: N = Sz(OAy)
If 0 > At Or At > N Then Err.Raise 1, , FmtQQ("At[?] is outside OAy-UB[?]", At, UB(OAy))
ReDim Preserve OAy(N)
Dim J&
For J = N To At + 1 Step -1
    OAy(J) = OAy(J - 1)
Next
OAy(At) = Ele
End Sub
Sub Push(O, P)
Dim N&: N = Sz(O)
ReDim Preserve O(N)
O(N) = P
End Sub
Sub RmvEle(Ay, Ele, Optional At& = -1)
Dim Idx&
    If At < 0 Then
        Idx = AyIdx(Ay, Ele)
    Else
        Idx = At
    End If
If Idx = -1 Then Exit Sub
Dim U&: U = UB(Ay)
If At > U Then Exit Sub
If U = 0 Then Erase Ay: Exit Sub
Dim J&
For J = Idx To U - 1
    Ay(J) = Ay(J + 1)
Next
ReDim Preserve Ay(U - 1)
End Sub

Function ChkEqAy(Ay1, Ay2, Optional Ay1Nm$ = "Exp", Optional Ay2Nm$ = "Act") As String()
Dim O$()
    If Sz(Ay1) <> Sz(Ay2) Then Push O, FmtQQ("Array [?] and [?] has different Sz: [?] [?]", Ay1Nm, Ay2Nm, Sz(Ay1), Sz(Ay2)): GoTo X
Dim O1$()
    Dim A2: A2 = Ay2
    Dim J%, ReachLimit As Boolean
    Dim V
    For Each V In Ay1
        If AyHas(A2, V) Then
            RmvEle A2, V
        Else
            Push O1, FmtQQ("Ele in [?] not found in {?]: [?]", Ay1Nm, Ay2Nm, V)
            J = J + 1
        End If
        If J > 10 Then
            ReachLimit = True
            Exit For
        End If
    Next
    If Not AyIsEmpty(A2) Then
        If J < 10 Then
            For Each V In A2
                Push O1, FmtQQ("Ele in [?] not found in {?]: [?]", Ay2Nm, Ay1Nm, V)
                J = J + 1
                If J > 10 Then
                    ReachLimit = True
                    Exit For
                End If
            Next
        End If
    End If
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
PushAy O, QuoteAy(Ay1, "[]")
Push O, FmtQQ("Ay-[?]:", Ay2Nm)
PushAy O, QuoteAy(Ay2, "[]")
ChkEqAy = O
End Function
Sub AssertEqAy(Ay1, Ay2, Optional Ay1Nm$ = "Exp", Optional Ay2Nm$ = "Act")
AssertChk ChkEqAy(Ay1, Ay2, Ay1Nm, Ay2Nm)
End Sub
Private Sub ChkEqAy__Tst()
AyDmp ChkEqAy(Array(1, 2, 3, 3, 4), Array(1, 2, 3, 4, 4))
End Sub
Sub AssertChk(Chk$())
If AyIsEmpty(Chk) Then Exit Sub
AyBrw Chk
Err.Raise 1, , "Error checked!"
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
Function Sz&(Ay)
On Error Resume Next
Sz = UBound(Ay) + 1
End Function
Function UB&(Ay)
UB = Sz(Ay) - 1
End Function
Function AyHas(Ay, Itm) As Boolean
If AyIsEmpty(Ay) Then Exit Function
Dim I
For Each I In Ay
    If I = Itm Then AyHas = True: Exit Function
Next
End Function
Sub RmvLasNEle(Ay, Optional NEle% = 1)
ReDim Preserve Ay(UB(Ay) - NEle)
End Sub
Function Pop(Ay)
Pop = AyLasEle(Ay)
RmvLasNEle Ay
End Function
Function AyLasEle(Ay)
AyLasEle = Ay(UB(Ay))
End Function
Function DblQuoteAy(Ay) As String()
Dim O$(), U&, J&
U = UB(Ay)
ReDim Preserve O(U)
For J = 0 To U
    O(J) = """" & Ay(J) & """"
Next
DblQuoteAy = O
End Function
Function AddAyPfx(Ay, Pfx) As String()
Dim O$(), U&, J&
U = UB(Ay)
If U = -1 Then Exit Function
ReDim Preserve O(U)
For J = 0 To U
    O(J) = Pfx & Ay(J)
Next
AddAyPfx = O
End Function
Function Sy(ParamArray Ap()) As String()
Dim Av(): Av = Ap
Sy = AySy(Av)
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
Function AyIdx&(Ay, Itm)
Dim J&
For J = 0 To UB(Ay)
    If Ay(J) = Itm Then AyIdx = J: Exit Function
Next
AyIdx = -1
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
Sub PushAy(OAy, Ay)
If AyIsEmpty(Ay) Then Exit Sub
Dim I
For Each I In Ay
    Push OAy, I
Next
End Sub
Function AyBrw(Ay)
Dim T$
T = TmpFt
WrtAy Ay, T
FtBrw T
End Function
Sub WrtAy(Ay, Ft)
WrtStr JnCrLf(Ay), Ft
End Sub
Function QuoteAy(Ay, QuoteStr$) As String()
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
QuoteAy = O
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
Private Sub AySrt__Tst()
Dim Exp, Act
Dim Ay
Ay = Array(1, 2, 3, 4, 5): Exp = Ay:                   Act = AySrt(Ay):       AssertEqAy Exp, Act
Ay = Array(1, 2, 3, 4, 5): Exp = Array(5, 4, 3, 2, 1): Act = AySrt(Ay, True): AssertEqAy Exp, Act
Ay = Array(":", "~", "P"): Exp = Array(":", "P", "~"): Act = AySrt(Ay):       AssertEqAy Exp, Act
'-----------------
Erase Ay
Push Ay, ":PjUpdTstFun:Sub"
Push Ay, ":SrcLinBrk:Function"
Push Ay, "~~:Tst:Sub"
Push Ay, ":PjTstFunNy_WithEr:Function"
Push Ay, "~Private:JnContinueLin__Tst:Sub"
Push Ay, "Private:IsPfx:Function"
Push Ay, "Private:MdFunDrs_FunBdyLy:Function"
Push Ay, "Private:MdFunEndLno:Function"
Erase Exp
Push Exp, ":PjTstFunNy_WithEr:Function"
Push Exp, ":PjUpdTstFun:Sub"
Push Exp, ":SrcLinBrk:Function"
Push Exp, "Private:IsPfx:Function"
Push Exp, "Private:MdFunDrs_FunBdyLy:Function"
Push Exp, "Private:MdFunEndLno:Function"
Push Exp, "~Private:JnContinueLin__Tst:Sub"
Push Exp, "~~:Tst:Sub"
Act = AySrt(Ay):       AssertEqAy Exp, Act
'---------------------
Ay = FtLy(Tst_ResPth & "AySrt_Ft1.txt")
Exp = FtLy(Tst_ResPth & "AySrt_Ft1_Exp.txt")
Act = AySrt(Ay):       AssertEqAy Exp, Act

End Sub

Function Tst_ResPth$()
Tst_ResPth = PjSrcPth & "TstRes\"
End Function
Private Sub Tst_ResPthBrw()
PthBrw Tst_ResPth
End Sub
Private Sub AySrtIntoIdxAy__Tst()
Dim Ay: Ay = Array("A", "B", "C", "D", "E")
AssertEqAy Array(0, 1, 2, 3, 4), AySrtIntoIdxAy(Ay)
AssertEqAy Array(4, 3, 2, 1, 0), AySrtIntoIdxAy(Ay, True)
End Sub
Function AySrtIntoIdxAy(Ay, Optional Des As Boolean) As Long()
If AyIsEmpty(Ay) Then AySrtIntoIdxAy = Ay: Exit Function
Dim Idx&, V, J&
Dim O&():
Push O, 0
For J = 1 To UB(Ay)
    AyIns O, J, AySrtIntoIdxAy_Idx(O, Ay, Ay(J), Des)
Next
AySrtIntoIdxAy = O
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
Sub AyDmp(Ay)
If AyIsEmpty(Ay) Then Exit Sub
Dim I
For Each I In Ay
    Debug.Print I
Next
End Sub
Sub Tst()
AySrt__Tst
AySrtIntoIdxAy__Tst
ChkEqAy__Tst
End Sub
