VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Public Ay_
'

Function Init(Ay) As Ay
If Not IsArray(Ay) Then Err.Raise 1, , FmtQQ("Given [Ay_] is not array, but TypeName[?]", TypeName(Ay_))
Ay_ = Ay
Set Init = Me
End Function
Function Brw() As Ay
Dim T$: T = Tmp.Ft
Wrt T
Ft(T).Brw
Set Brw = Me
End Function
Function RmvLasChr() As String()
If IsEmpty Then Exit Function
Dim O$()
ReDim O(UB)
Dim J&
For J = 0 To UB
    O(J) = Ay_(J)
Next
End Function
Sub PutAt(At As Range)
HSq.PutAt At
End Sub
Function HSq() As Sq
Set HSq = SqByHAy(Ay_)
End Function
Function VSq() As Sq
Set HSq = SqByVAy(Ay_)
End Function

Sub AsrtChk(Chk$())
If AyIsEmpty(Chk) Then Exit Sub
Ay(Chk).Brw
Err.Raise 1, , "Error checked!"
End Sub

Sub AsrtEq(Ay2, Optional Ay1Nm$ = "Exp", Optional Ay2Nm$ = "Act")
AsrtChk ChkEq(Ay2, Ay1Nm, Ay2Nm)
End Sub
Function Clone() As Ay
Set Clone = Nw.Ay(Ay_)
End Function
Function AddAy(ParamArray AyAp()) As Ay
Dim Av(): Av = AyAp
Dim A
Dim O As Ay: Set O = Clone
For Each A In Av
    O.PushAy A
Next
Set AddAy = O
End Function

Function AddPfx(Pfx) As String()
Dim O$(), U&, J&
U = UB
If U = -1 Then Exit Function
ReDim Preserve O(U)
For J = 0 To U
    O(J) = Pfx & Ay_(J)
Next
AddPfx = O
End Function
Function SliceCnt(Idx&, Cnt&)
SliceCnt = Slice(Idx, Idx + Cnt - 1)
End Function
Function SliceCntX(Idx&, Cnt&) As Ay
Set SliceCntX = Ay(SliceCnt(Idx, Cnt))
End Function
Sub Asg(ParamArray OAp())
Dim Av(): Av = OAp
Dim J%
For J = 0 To Vb.UB(Av)
    If Not IsMissing(OAp(J)) Then
         Vb.Asg Ay(J), OAp(J)
    End If
Next
End Sub
Function Slice(FmIdx&, ToIdx&)
Dim O: O = Ay_: Erase O
Dim J&
For J = FmIdx To ToIdx
    Vb.Push O, Ay_(J)
Next
Slice = O
End Function
Function SliceX(FmIdx&, ToIdx&) As Ay
Set SliceX = Ay(Slice(FmIdx, ToIdx))
End Function
Sub Asg_Idx(IdxAy&(), ParamArray OAp())
Dim J%
For J = 0 To Vb.UB(IdxAy)
    OAp(J) = Ay_(IdxAy(J))
Next
End Sub
 Function Has(Itm) As Boolean
If IsEmpty Then Exit Function
Dim I
For Each I In Ay_
    If I = Itm Then Has = True: Exit Function
Next
End Function

Function Idx&(Itm)
Dim J&
For J = 0 To UB
    If Ay_(J) = Itm Then Idx = J: Exit Function
Next
Idx = -1
End Function

Function IdxAy(SubAy) As Long()
If AyIsEmpty(SubAy) Then Exit Function
Dim O&()
Dim U&: U = Vb.UB(SubAy)
ReDim O(U)
Dim J&
For J = 0 To U
    O(J) = Idx(SubAy(J))
Next
IdxAy = O
End Function

Function Ins(Optional Ele = Empty, Optional At&) As Ay
Dim N&: N = Sz
If 0 > At Or At > N Then Err.Raise 1, , FmtQQ("At[?] is outside OAy-UB[?]", At, UB)
ReDim Preserve Ay_(N)
Dim J&
For J = N To At + 1 Step -1
    Ay_(J) = Ay_(J - 1)
Next
Ay_(At) = Ele
Set Ins = Me
End Function

Function LasEle()
Vb.Asg Ay_(UB), LasEle
End Function

Function Minus(ParamArray AyAp())
Dim Av(): Av = AyAp
Dim IAy
Dim O As Ay: O = Clone
For Each IAy In Av
    If O.IsEmpty Then Set Minus = O: Exit Function
    O.RmvEleAy IAy
Next
Set Minus = O
End Function
Private Sub RmvEmptyEle__Tst()
Ay_ = Array(1, 2, Empty, "", 1, "", "", 3)
Dim Act: Act = RmvEmptyEle.Ay_
Debug.Assert Vb.Sz(Act) = 4
Debug.Assert Act(0) = 1
Debug.Assert Act(1) = 2
Debug.Assert Act(2) = 1
Debug.Assert Act(3) = 3
End Sub
Function RmvEmptyEle() As Ay
Dim J&, I&
If IsEmpty Then Set RmvEmptyEle = Me: Exit Function
If Vb.IsEmpty(Ay_(0)) Then
    I = 0
Else
    I = 1
End If
For J = 1 To UB
    If Vb.IsEmpty(Ay_(J)) Then
    Else
        If J > I Then
            Vb.Asg Ay_(J), Ay_(I)
        ElseIf J < I Then Stop
        End If
        I = I + 1
    End If
Next
If I = 0 Then
    Erase Ay_
Else
    ReDim Preserve Ay_(I - 1)
End If
Set RmvEmptyEle = Me
End Function

Function RmvEleAy(Ay) As Ay
If Not IsEmpty Then
Dim V
If Not AyIsEmpty(Ay) Then
For Each V In Ay
    RmvEle V
Next
End If
End If
Set RmvEleAy = Me
End Function

Function RmvEle(Ele) As Ay
Dim I&
I = Idx(Ele)
If I >= 0 Then RmvAt I
Set RmvEle = Me
End Function
Function RmvAt(Optional At& = 0) As Ay
Set RmvAt = RmvAtCnt(At)
End Function

Function RmvAtCnt(At&, Optional Cnt& = 1) As Ay
If Cnt <= 0 Then GoTo X
If At < 0 Then GoTo X
Dim U&: U = UB
If At > U Then GoTo X
If U = 0 Then GoTo X
Dim J&
For J = At To U - Cnt
    Ay_(J) = Ay_(J + Cnt)
Next
ReDim Preserve Ay_(U - Cnt)
X:
Set RmvAtCnt = Me
End Function

Function RmvFstEle() As Ay
Set RmvFstEle = RmvAt
End Function

Function RmvLasEle() As Ay
Set RmvLasEle = RmvAt(UB)
End Function
Function Clr() As Ay
Erase Ay_
Set Clr = Me
End Function
Function SelX(IdxAy&()) As Ay
Set SelX = Nw.Ay(Sel(IdxAy))
End Function
Function Sel(IdxAy&())
Dim O: O = Ay_: Erase O
If Not AyIsEmpty(IdxAy) Then
    Dim U&: U = Vb.UB(IdxAy)
    ReDim O(U)
    Dim J&
    For J = 0 To U
        Asg Ay_(J), O(J)
    Next
End If
Sel = O
End Function
Function Srt(Optional Des As Boolean) As Ay
If IsEmpty Then Set Srt = Me: Exit Function
Dim Idx&, V, J&
Dim O As Ay: Set O = Clone.Clr
O.Push Ay_(0)
For J = 1 To UB
    O.Ins Ay_(J), O.SrtIdx(Ay(J), Des)
Next
Set Srt = O
End Function
Function SrtIdx&(V, Des As Boolean)
Dim I, O&
If Des Then
    For Each I In Ay_
        If V > I Then SrtIdx = O: Exit Function
        O = O + 1
    Next
    SrtIdx = O
    Exit Function
End If
For Each I In Ay_
    If V < I Then SrtIdx = O: Exit Function
    O = O + 1
Next
SrtIdx = O
End Function
Private Sub SrtIntoIdxAy__Tst()
Ay_ = Array("a", "c", "d", "e", "h", "f")
Dim Act&(): Act = SrtIntoIdxAy
Debug.Assert Vb.Sz(Ay_) = Vb.Sz(Act)
Dim J&
For J = 1 To Vb.UB(Act)
    If Ay_(Act(J - 1)) > Ay_(Act(J)) Then Stop
Next
End Sub

Function SrtIntoIdxAy(Optional Des As Boolean) As Long()
If IsEmpty Then Exit Function
Dim Ele&, Idx&(), At&
Dim O As Ay: Set O = EmptyLngAy
O.Push 0
For Ele = 1 To UB
    Idx = O.Ay_
    At = SrtIntoIdxAy_At(Idx, Ay_, Ay_(Ele), Des)
    O.Ins Ele, At
Next
SrtIntoIdxAy = O.Ay_
End Function
Private Function SrtIntoIdxAy_At&(Idx&(), Ay, V, Des As Boolean)
'Ay(Idx(0)), .., Ay(Idx(U)) is sorted
'It is required to find V should be inserted in which Idx within 0..U
'{Ay} is the whole to be sorted.  No sorting is done in {Ay}, only done in Idx&
Dim I, O&
If Des Then
    For Each I In Idx
        If V > Ay(I) Then SrtIntoIdxAy_At& = O: Exit Function
        O = O + 1
    Next
    SrtIntoIdxAy_At& = O
    Exit Function
End If
For Each I In Idx
    If V < Ay(I) Then SrtIntoIdxAy_At& = O: Exit Function
    O = O + 1
Next
SrtIntoIdxAy_At& = O
End Function

Function Trim() As Ay
Dim J&
For J = 0 To UB
    Ay_(J) = VBA.Trim(Ay_(J))
Next
Set Trim = Me
End Function
Function Lines$()
Lines = JnCrLf
End Function
Sub Wrt(Ft)
StrWrt JnCrLf, Ft
End Sub

Function ChkEq(Ay, Optional Ay1Nm$ = "Exp", Optional Ay2Nm$ = "Act") As String()
Dim U&: U = Vb.UB(Ay)
Dim O$()
    If U <> UB Then Vb.Push O, FmtQQ("Array [?] and [?] has different Sz: [?] [?]", Ay1Nm, Ay2Nm, Sz, Vb.Sz(Ay)): GoTo X
If IsEmpty Then Exit Function
Dim O1$()
    Dim A2: A2 = Ay
    Dim J&, ReachLimit As Boolean
    Dim Cnt%
    For J = 0 To U
        If Ay_(J) <> Ay(J) Then
            Vb.Push O1, FmtQQ("[?]-th Ele is diff: ?[?]<>?[?]", Ay1Nm, Ay2Nm, Ay_(J), Ay(J))
            Cnt = Cnt + 1
        End If
        If Cnt > 10 Then
            ReachLimit = True
            Exit For
        End If
    Next
If Vb.AyIsEmpty(O1) Then Exit Function
Dim O2$()
    Vb.Push O2, FmtQQ("Array [?] and [?] both having size[?] have differnt element(s):", Ay1Nm, Ay2Nm, Sz)
    If ReachLimit Then
        Vb.Push O2, FmtQQ("At least [?] differences:", Vb.Sz(O1))
    End If
Vb.PushAy O, O2
Vb.PushAy O, O1
X:
Vb.Push O, FmtQQ("Ay-[?]:", Ay1Nm)
Vb.PushAy O, Clone.Quote("[]")
Vb.Push O, FmtQQ("Ay-[?]:", Ay2Nm)
Vb.PushAy O, Nw.Ay(Ay).Quote("[]")
ChkEq = O
End Function
Function DblQuoteJnComma$()
DblQuoteJnComma = Nw.Ay(DblQuote).JnComma
End Function
Function DblQuote() As String()
Dim O$(), U&, J&
U = UB
ReDim Preserve O(U)
For J = 0 To U
    O(J) = """" & Ay(J) & """"
Next
DblQuote = O
End Function
Function CsvQQStr$()
If IsEmpty Then Exit Function
Dim O$()
Dim I, J%
ReDim O(UB)
For Each I In Ay_
    O(J) = VbTyInsSqlValPhraseTp(VarType(I))
    J = J + 1
Next
CsvQQStr = Vb.JnComma(O)
End Function

Function Pop()
Vb.Asg LasEle, Pop
RmvLasEle
End Function

Function Push(I) As Ay
Dim N&: N = Sz
ReDim Preserve Ay_(N)
Vb.Asg I, Ay_(N)
Set Push = Me
End Function
Private Sub RmvEmptyEleAtEnd__Tst()
Dim Ay As Ay: Set Ay = Ay = Ay(Array(Empty, Empty, Empty, 1, Empty, Empty))
Ay.RmvEmptyEleAtEnd:  Debug.Assert Ay.Sz = 4
End Sub
Function RmvEmptyEleAtEnd() As Ay
Dim LasU&, U&
For LasU = UB To 0 Step -1
    If Not Vb.IsEmpty(Ay_(LasU)) Then
        Exit For
    End If
Next
If LasU = -1 Then
    Erase Ay_
Else
    ReDim Preserve Ay_(LasU)
End If
Set RmvEmptyEleAtEnd = Me
End Function
Function IsEq(Ay) As Boolean
Dim U&: U = UB: If U <> Vb.UB(Ay) Then Exit Function
Dim J&
For J = 0 To U
    If Ay_(J) <> Ay(J) Then Exit Function
Next
IsEq = True
End Function
Function PushAy(Ay) As Ay
If Not IsEmpty Then
    Dim I
    For Each I In Ay
        Push I
    Next
End If
Set PushAy = Me
End Function

Sub PushNonEmpty(I)
If Vb.IsEmpty(I) Then Exit Sub
Push I
End Sub
Function Quote(QuoteStr$) As String()
If IsEmpty Then Exit Function
Dim U&: U = UB
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
Quote = O
End Function

Function RmvLasNEle(Optional NEle% = 1) As Ay
If NEle = Sz Then Set RmvLasNEle = Clr
ReDim Preserve Ay_(UB - NEle)
End Function

Function Sz&()
On Error Resume Next
Sz = UBound(Ay_) + 1
End Function

Function JnComma$()
JnComma = Join(Ay_, ",")
End Function

Function JnCrLf()
JnCrLf = Join(Ay_, vbCrLf)
End Function

Function JnSpc()
JnSpc = Join(Ay_, " ")
End Function

Private Sub Tst_ResPthBrw()
Pth(Tst_ResPth).Brw
End Sub

Sub AddAy__Tst()
Dim Act(), Exp(), Ay1 As Ay, Ay2()
Set Ay1 = Nw.Ay(Array(1, 2, 2, 2, 4, 5))
Ay2 = Array(2, 2)
Act = Ay_.AddAy(Ay2)
Exp = Array(1, 2, 2, 2, 4, 5, 2, 2)
Ay_.AsrtEq Exp
Ay_.AsrtEq Array(1, 2, 2, 2, 4, 5)
Ay_.AsrtEq Array(2, 2)
End Sub

Sub AyMinusOneAy__Tst()
Dim Act(), Exp()
Dim Ay1(), Ay2()
Ay1 = Array(1, 2, 2, 2, 4, 5)
Ay2 = Array(2, 2)
'Act = AyMinusOneAy(Ay1, Ay2)
Exp = Array(1, 2, 4, 5)
'AsrtEqAy Exp, Act
'
'Act = AyMinus(Array(1, 2, 2, 2, 4, 5), Array(2, 2), Array(5))
'Exp = Array(1, 2, 4)
'AsrtEqAy Exp, Act
End Sub

Private Sub AyRmvEleAtCnt__Tst()
Dim Ay(): Ay = Array(1, 2, 3, 4, 5)
'AyRmvEleAtCnt Ay, 1, 2
'AsrtEqAy Array(1, 4, 5), Ay
End Sub

Private Sub AySrt__Tst()
'Dim Exp, Act
'Dim Ay
'Ay = Array(1, 2, 3, 4, 5): Exp = Ay:                   Act = AySrt(Ay):       AsrtEqAy Exp, Act
'Ay = Array(1, 2, 3, 4, 5): Exp = Array(5, 4, 3, 2, 1): Act = AySrt(Ay, True): AsrtEqAy Exp, Act
'Ay = Array(":", "~", "P"): Exp = Array(":", "P", "~"): Act = AySrt(Ay):       AsrtEqAy Exp, Act
''-----------------
'Erase Ay
'Push Ay, ":PjUpdTstFun:Sub"
'Push Ay, ":SrcLinBrk:Function"
'Push Ay, "~~:Tst:Sub"
'Push Ay, ":PjTstFunNy_WithEr:Function"
'Push Ay, "~Private:JnContinueLin__Tst:Sub"
'Push Ay, "Private:IsPfx:Function"
'Push Ay, "Private:MdFunDrs_FunBdyLy:Function"
'Push Ay, "Private:MdFunEndLno:Function"
'Erase Exp
'Push Exp, ":PjTstFunNy_WithEr:Function"
'Push Exp, ":PjUpdTstFun:Sub"
'Push Exp, ":SrcLinBrk:Function"
'Push Exp, "Private:IsPfx:Function"
'Push Exp, "Private:MdFunDrs_FunBdyLy:Function"
'Push Exp, "Private:MdFunEndLno:Function"
'Push Exp, "~Private:JnContinueLin__Tst:Sub"
'Push Exp, "~~:Tst:Sub"
'Act = AySrt(Ay):       AsrtEqAy Exp, Act
''---------------------
'Ay = FtLy(Tst_ResPth & "AySrt_Ft1.txt")
'Exp = FtLy(Tst_ResPth & "AySrt_Ft1_Exp.txt")
'Act = AySrt(Ay):       AsrtEqAy Exp, Act

End Sub
Function AySel(Ay, IdxAy&())
'Dim U&
'    U = UB(IdxAy)
'Dim O
'    O = Ay
'    ReDim O(U)
'Dim J&
'For J = 0 To U
'    O(J) = Ay(IdxAy(J))
'Next
'AySel = O
End Function


Private Sub AySrtIntoIdxAy__Tst()
'Dim Ay: Ay = Array("A", "B", "C", "D", "E")
'AsrtEqAy Array(0, 1, 2, 3, 4), AySrtIntoIdxAy(Ay)
'AsrtEqAy Array(4, 3, 2, 1, 0), AySrtIntoIdxAy(Ay, True)
End Sub

Private Sub ChkEqAy__Tst()
'AyDmp ChkEqAy(Array(1, 2, 3, 3, 4), Array(1, 2, 3, 4, 4))
End Sub

Sub Tst()
RmvEmptyEle__Tst
SrtIntoIdxAy__Tst
End Sub

Function Dmp() As Ay
Debug.Print "Sz=" & Sz
Dim J&
For J = 0 To UB
    Debug.Print J; " [" & Ay_(J) & "]"
Next
Set Dmp = Me
End Function
Function Itm(I&)
Asg Ay_(I), Itm
End Function
Function StrAy() As String()
Dim O$()
If Not IsEmpty Then
    ReDim O(UB)
    Dim J&, I
    For Each I In Ay_
        O(J) = I
    Next
End If
StrAy = O
End Function
Function IsObjAy() As Boolean
IsObjAy = VarType(Ay_) And vbObject
End Function
Function IsStrAy() As Boolean
IsStrAy = VarType(Ay_) And vbString
End Function

Function IsEmpty() As Boolean
IsEmpty = Sz = 0
End Function
Function UB&()
UB = Sz - 1
End Function
