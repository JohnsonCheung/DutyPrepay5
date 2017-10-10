Attribute VB_Name = "SqTp"
Option Compare Database
Option Explicit
'Sqpi:=Sql-Phrase-Item
Private Const SqpiLvs1$ = "|Sel selDis gp into fm left jn" ' Sqpi = Sql-Phrase-Itm
Private Const SqpiLvs2$ = " whBetStr whBetNbr whInStrLis whInNbrLis"
Private Const SqpiLvs3$ = " andBetStr andBetNbr andInStrLis andInNbrLis"
Public Const SqpiLvs$ = SqpiLvs1 & SqpiLvs2 & SqpiLvs3
Public Enum eBlkTy
    eBlkRmk
    eBlkPrm
    eBlkSw
    eBlkSq
    eBlkEr
End Enum
Public Type Blk
    Ty As eBlkTy
    Ly() As String
    FstLin As String
End Type

Function A_Main(TpLines$) As StrOpt
A_Main = Evl(TpLines$)
End Function

Function Blk(I, IsFstBlk As Boolean) As Blk
Dim Ly$()
Dim FstLin$
    Ly = SplitCrLf(I)
    If IsFstBlk Then
        FstLin = ""
    Else
        If Sz(Ly) > 0 Then FstLin = Ly(0)
        Ly = AyRmvEleAt(Ly)
    End If

Blk.Ty = BlkTy(Ly)
Blk.Ly = AyRTrim(AyRmv2Dash(Ly))
Blk.FstLin = FstLin
End Function

Function BlkAyDic(A() As Blk, Ty As eBlkTy) As Dictionary
Set BlkAyDic = DicByLy(BlkAySelLy(A, Ty), IgnoreDup:=True)
End Function

Function BlkAyDicPrm(A() As Blk) As Dictionary
Set BlkAyDicPrm = BlkAyDic(A, eBlkPrm)
End Function

Function BlkAyDicSw(A() As Blk) As Dictionary
Set BlkAyDicSw = BlkAyDic(A, eBlkSw)
End Function

Function BlkAySelLy(A() As Blk, Ty As eBlkTy) As String()
Const CSub$ = "BlkAySelOne"
Dim O$()
Dim J%
For J = 0 To BlkUB(A)
    If A(J).Ty = Ty Then BlkAySelLy = A(J).Ly: Exit Function
Next
Er CSub, "There is no such {Ty} for given {N-BlkAy}", BlkTyToStr(Ty), BlkSz(A)
End Function

Function BlkAySelLyAy(A() As Blk, Ty As eBlkTy) As Variant()
Dim O()
Dim J%
For J = 0 To BlkUB(A)
If A(J).Ty = Ty Then Push O, A(J).Ly
Next
BlkAySelLyAy = O
End Function

Function BlkBrk(TpLines$) As Blk()
Dim A$(): A = Split(TpLines, vbCrLf & "==")
If AyIsEmpty(A) Then Exit Function
Dim O() As Blk, I, J%
J = 0
For Each I In A
    BlkPush O, Blk(I, IsFstBlk:=(J = 0))
    J = J + 1
Next
BlkBrk = O
End Function

Sub BlkDmp(A As Blk)
AyDmp BlkLy(A)
End Sub

Function BlkLy(A As Blk) As String()
Dim O$()
Push O, BlkTyToStr(A.Ty)
PushAy O, A.Ly
BlkLy = O
End Function

Function BlkLyFstLin$(Ly$())
Dim J%
For J = 0 To UB(Ly)
    If Not IsRmkLin(Ly(J)) Then BlkLyFstLin = Ly(J): Exit Function
Next
End Function

Function BlkLyIsPrm(Ly$()) As Boolean
If AyIsEmpty(Ly) Then Exit Function
Dim L$: L = BlkLyFstLin(Ly)
BlkLyIsPrm = FstChr(L) = "%"
End Function

Function BlkLyIsRmk(Ly$()) As Boolean
If AyIsEmpty(Ly) Then BlkLyIsRmk = True: Exit Function
End Function

Function BlkLyIsSq(Ly$()) As Boolean
If AyIsEmpty(Ly) Then Exit Function
Dim Ay(): Ay = Array("?sel ", "sel ", "?seldis ", "seldis", "upd ", "drp ")
Dim L$: L = BlkLyFstLin(Ly)
If IsPfxAy(L, Ay) Then BlkLyIsSq = True: Exit Function
End Function

Function BlkLyIsSw(Ly$()) As Boolean
If AyIsEmpty(Ly) Then Exit Function
Dim L$: L = BlkLyFstLin(Ly)
If FstChr(L) = "?" Then
    If BlkLyIsSq(Ly) Then Exit Function
    BlkLyIsSw = True: Exit Function
End If
End Function

Sub BlkPush(O() As Blk, I As Blk)
Dim N%: N = BlkSz(O)
ReDim Preserve O(N)
O(N) = I
End Sub

Function BlkSz%(A() As Blk)
On Error Resume Next
BlkSz = UBound(A) + 1
End Function

Function BlkTy(Ly$()) As eBlkTy
If BlkLyIsPrm(Ly) Then BlkTy = eBlkPrm: Exit Function
If BlkLyIsSw(Ly) Then BlkTy = eBlkSw: Exit Function
If BlkLyIsSq(Ly) Then BlkTy = eBlkSq: Exit Function
If BlkLyIsRmk(Ly) Then BlkTy = eBlkRmk: Exit Function
BlkTy = eBlkEr
End Function

Function BlkTyToStr$(A As eBlkTy)
Dim O$
Select Case A
Case eBlkEr: O = "*ErBlk"
Case eBlkPrm: O = "*PrmBlk"
Case eBlkSw: O = "*SwBlk"
Case eBlkSq: O = "*SqBlk"
Case eBlkRmk: O = "*RmkBlk"
Case Else: Stop
End Select
BlkTyToStr = O
End Function

Function BlkUB%(A() As Blk)
BlkUB = BlkSz(A) - 1
End Function

Function ChkBlk(A As Blk, Prm As Dictionary, Sw As Dictionary) As ChkRslt()
Dim O() As ChkRslt
    With A
        Select Case .Ty
        Case eBlkEr
            Stop
        Case eBlkPrm: O = ChkPrm(A.Ly)
        Case eBlkSw:  O = ChkSw(A.Ly, Prm, Sw)
        Case eBlkSq:  O = ChkSq(A.Ly, Prm, Sw)
        Case eBlkRmk
        Case Else: Stop
        End Select
    End With
ChkBlk = O
End Function

Function ChkBlkAy(A() As Blk) As StrOpt
'ChkBlkAy:=Check-Block-Array
Dim BlkLines$()
Dim IsEr%
    Dim R() As ChkRslt
    Dim Ay$()
    Dim J%
    Dim Prm As Dictionary
    Dim Sw As Dictionary
        Set Prm = BlkAyDicPrm(A)
        Set Sw = BlkAyDicSw(A)
    For J = 0 To BlkUB(A)
        R = ChkBlk(A(J), Prm, Sw)
        Ay = ChkRsltPut(A(J).Ly, R)
        Push BlkLines, JnCrLf(Ay)
        If ChkRsltSz(R) > 0 Then IsEr = True
    Next
If IsEr Then ChkBlkAy = SomStr(Join(Ay, vbCrLf & "=="))
End Function

Function ChkBlkLyDupKey(Ly$()) As ChkRslt()
':ChkBlkLyDupKey=Check-Ly-DupKey
Dim Ky$()
    Ky = ChkBlkLyDupKey_Ky(Ly)
    Ky = AyDupAy(Ky)
Dim O() As ChkRslt
    Dim J%
    For J = 0 To UB(Ly)
        If ChkBlkLyDupKey_IsDup(Ly(J), Ky) Then ChkRsltPush O, ChkRslt("Key is dup", J)
    Next
ChkBlkLyDupKey = O
End Function

Function ChkBlkLyDupKey_IsDup(Lin, Ky$()) As Boolean
Dim Fst$: Fst = FstTerm(Lin)
ChkBlkLyDupKey_IsDup = AyHas(Ky, Fst)
End Function

Function ChkBlkLyDupKey_Ky(Ly$()) As String()
Dim J%, O$()
For J = 0 To UB(Ly)
    If Not IsRmkLin(Ly(J)) Then
        Push O, FstTerm(Ly(J))
    End If
Next
ChkBlkLyDupKey_Ky = O
End Function

Function ChkPrm(PrmLy$()) As ChkRslt()
Dim O() As ChkRslt
    ChkRsltPushAy O, ChkPrm__ForEachLine(PrmLy)
    ChkRsltPushAy O, ChkPrm__ForDupKey(PrmLy)
ChkPrm = O
End Function

Function ChkPrm__ForDupKey(PrmLy$()) As ChkRslt()
Dim O() As ChkRslt
    ChkRsltPushAy O, ChkBlkLyDupKey(PrmLy)
ChkPrm__ForDupKey = O
End Function

Function ChkPrm__ForEachLine(PrmLy$()) As ChkRslt()
Dim J%, A$, O() As ChkRslt
For J = 0 To UB(PrmLy)
    With ChkPrm_ForOneLine(PrmLy(J))
        If .Som Then ChkRsltPush O, ChkRslt(.Str, J)
    End With
Next
ChkPrm__ForEachLine = O
End Function

Function ChkPrm_ForOneLine(PrmLin) As StrOpt
Dim PrmNm$, Rst$
    Dim L$
    L = PrmLin
    PrmNm = ParseTerm(L)
    Rst = L
Dim A$
A = ChkPrm_ForPrmNmPfx(PrmNm):     If A <> "" Then ChkPrm_ForOneLine = SomStr(A): Exit Function
A = ChkPrm_ForPrmIsSw(PrmNm, Rst): If A <> "" Then ChkPrm_ForOneLine = SomStr(A): Exit Function
End Function

Function ChkPrm_ForPrmIsSw$(PrmNm$, Rst$)
If Not IsPfx(PrmNm, "%?") Then Exit Function
Dim Ay$()
Ay = SplitLvs(Rst)
If Sz(Ay) <> 1 Then
    ChkPrm_ForPrmIsSw = "For %?xxx, it must have 2 terms": Exit Function
End If
If Ay(0) <> "0" And Ay(0) <> "1" Then ChkPrm_ForPrmIsSw = "For %?xxx, 2nd term must be 1 or 0": Exit Function
End Function

Function ChkPrm_ForPrmNmPfx$(PrmNm$)

End Function

Function ChkSq(SqLy$(), Prm As Dictionary, Sw As Dictionary) As ChkRslt()
Dim O() As ChkRslt
    ChkRsltPushAy O, ChkSq__ForEachLine(SqLy, Prm, Sw)
    ChkRsltPushAy O, ChkSq__EndMsg(SqLy)
ChkSq = O
End Function

Function ChkSq__EndMsg(SqLy$()) As ChkRslt()
Dim O$()
PushNonEmpty O, ChkSq_SelMustHaveInto(SqLy)
PushNonEmpty O, ChkSq_SelMustHaveFm(SqLy)
If HasSubStr(Join(SqLy), "Invalid Sql-Phrase-Item") Then
    Push O, "Valid sql-phrase-item are: " & SqpiLvs
End If
ChkSq__EndMsg = ChkRsltMkEndMsg(O)
End Function

Function ChkSq__ForEachLine(SqLy$(), Prm As Dictionary, Sw As Dictionary) As ChkRslt()
':Cly=Check-Ly-Sq
Dim Expr As Dictionary
    Set Expr = ExprDic(SqLy)
Dim O() As ChkRslt
    Dim J%
    For J = 0 To UB(SqLy)
        With ChkSq_ForOneLine(SqLy(J), Prm, Sw, Expr)
            If .Som Then ChkRsltPush O, ChkRslt(.Str, J)
        End With
    Next
ChkSq__ForEachLine = O
End Function

Function ChkSq_ForOneLine(SqLin$, Prm As Dictionary, Sw As Dictionary, Expr As Dictionary) As StrOpt
'Clsq:=Check-Line-Sql
If IsRmkLin(SqLin) Then Exit Function
Dim L$
    L = SqLin
Dim Sqpi$ ' Sql-Phrase-Str
    Sqpi = RmvPfx(ParseTerm(L), "?")
Dim O$
    Select Case Sqpi
    Case "Gp", "Sel", "SelDis", "AndInStrLis", "AndInNbrLis"
    Case "WhBetStr"
    Case "Into", "Fm", "Upd"
    Case Else: O = "Invalid Sql-Phrase-Item"
    End Select
If O <> "" Then ChkSq_ForOneLine = SomStr(O)
End Function

Function ChkSq_SelMustHaveFm$(SqLy$())

End Function

Function ChkSq_SelMustHaveInto$(SqLy$())

End Function

Function ChkSw(SwLy$(), Prm As Dictionary, Sw As Dictionary) As ChkRslt()
Dim O() As ChkRslt
    ChkRsltPushAy O, ChkSw__ForDupKey(SwLy)
    ChkRsltPushAy O, ChkSw__ForEachLine(SwLy, Prm, Sw)
ChkSw = O
End Function

Function ChkSw__ForDupKey(SwLy$()) As ChkRslt()
ChkSw__ForDupKey = ChkBlkLyDupKey(SwLy)
End Function

Function ChkSw__ForEachLine(SwLy$(), Prm As Dictionary, Sw As Dictionary) As ChkRslt()
Dim J%, O() As ChkRslt
For J = 0 To UB(SwLy)
    With ChkSw_ForOneLin(SwLy(J), Prm, Sw)
        If .Som Then ChkRsltPush O, ChkRslt(.Str, J)
    End With
Next
ChkSw__ForEachLine = O
End Function

Function ChkSw_ForOneLin(SwLin, Prm As Dictionary, Sw As Dictionary) As StrOpt
'Clsw:=Check-Line-Switch
If IsRmkLin(SwLin) Then Exit Function
Dim L$
    L = SwLin
Dim NTerm%, Op$, TermAy$(), SwNm$
    SwNm = ParseTerm(L)
    Op = ParseTerm(L)
    TermAy = SplitLvs(L)
    NTerm = Sz(TermAy)
Dim A$
A = ChkSw_ForSwNmPfx(SwNm):                              If A <> "" Then ChkSw_ForOneLin = SomStr(A): Exit Function
A = ChkSw_ForOp(Op):                                     If A <> "" Then ChkSw_ForOneLin = SomStr(A): Exit Function
A = ChkSw_ForOnly2Term_For_EQ_and_NE(Op, NTerm):         If A <> "" Then ChkSw_ForOneLin = SomStr(A): Exit Function
A = ChkSw_ForTermAy_For_EQ_and_NE(Op, TermAy, Prm, Sw):  If A <> "" Then ChkSw_ForOneLin = SomStr(A): Exit Function
A = ChkSw_ForTermAy_For_AND_and_OR(Op, TermAy, Prm, Sw): If A <> "" Then ChkSw_ForOneLin = SomStr(A): Exit Function
End Function

Function ChkSw_ForOnly2Term_For_EQ_and_NE$(Op$, NTerm%)
If Op <> "EQ" And Op <> "NE" Then Exit Function
If NTerm = 2 Then Exit Function
ChkSw_ForOnly2Term_For_EQ_and_NE = "When 2nd-Term (Operator) is [AND OR], only 2 terms are allowed"
End Function

Function ChkSw_ForOp$(Op$)
If AyHas(Array("NE", "EQ", "AND", "OR"), Op) Then Exit Function
ChkSw_ForOp = "Operation in 2nd-Term must be [EQ NE AND OR]"
End Function

Function ChkSw_ForSwNmPfx$(SwNm$)
If Not IsPfx(SwNm, "?") Then ChkSw_ForSwNmPfx = "Must begin with ?"
End Function

Function ChkSw_ForTermAy_For_AND_and_OR$(Op$, TermAy$(), Prm As Dictionary, Sw As Dictionary)
If Op <> "AND" And Op <> "OR" Then Exit Function
If AyIsEmpty(TermAy) Then ChkSw_ForTermAy_For_AND_and_OR = "Must at one term after OR|AND": Exit Function
Dim O0$(), O1$(), O2$(), I
For Each I In TermAy
    If IsPfx(I, "?") Then
        If Not Sw.Exists(I) Then Push O0, I
    ElseIf IsPfx(I, "%?") Then
        If Not Prm.Exists(I) Then Push O1, I
    Else
        Push O2, I
    End If
Next
Dim A$, B$
If Not AyIsEmpty(O0) Then A = FmtQQ("[?] must be found in Switch", JnSpc(O0))
If Not AyIsEmpty(O1) Then A = FmtQQ("[?] must be found in Prm", JnSpc(O1))
If Not AyIsEmpty(O2) Then B = FmtQQ("[?] must begin with [ ? | %? ]", JnSpc(O1))
Dim O$()
    PushNonEmpty O, A
    PushNonEmpty O, B
ChkSw_ForTermAy_For_AND_and_OR = JnSpc(O)
End Function

Function ChkSw_ForTermAy_For_EQ_and_NE$(Op$, TermAy$(), Prm As Dictionary, Sw As Dictionary)
If Op <> "EQ" And Op <> "NE" Then Exit Function
If Sz(TermAy) <> 2 Then ChkSw_ForTermAy_For_EQ_and_NE = "For OR|AND, must have 2 operands": Exit Function
Select Case FstChr(TermAy(0))
    Case "%"
        If Not Prm.Exists(TermAy(0)) Then ChkSw_ForTermAy_For_EQ_and_NE = "For OR|AND, first term must be found in Prm": Exit Function
    Case Else
        ChkSw_ForTermAy_For_EQ_and_NE = "For OR|AND, first operand must begin with %": Exit Function
    End Select
    
Select Case FstChr(TermAy(1))
    Case "%"
        If Not Prm.Exists(TermAy()) Then
            ChkSw_ForTermAy_For_EQ_and_NE = "For EQ|NE, second operand not found in Prm": Exit Function
        End If
    Case "?"
        ChkSw_ForTermAy_For_EQ_and_NE = "For EQ|NE, second operand cannot begin with ?": Exit Function
    Case "*"
        If UCase(TermAy(1)) <> "*BLANK" Then
            ChkSw_ForTermAy_For_EQ_and_NE = "For EQ|NE, second operand can be *BLANK, but nothing else begin with *": Exit Function
        End If
    End Select
End Function

Function Evl(TpLines$) As StrOpt
Dim A() As Blk
    A = BlkBrk(TpLines)
Dim C As StrOpt
    C = ChkBlkAy(A)
If C.Som Then
    Evl = C
Else
    Evl = SomStr(EvlBlkAy(A))
End If
End Function

Function EvlBlkAy$(A() As Blk)
Dim Prm As Dictionary
Dim Sw As Dictionary
Set Prm = EvlBlkAyPrm(A)
Set Sw = EvlBlkAySw(A, Prm)
EvlBlkAy = EvlBlkAySq(A, Prm, Sw)
End Function

Function EvlBlkAyPrm(A() As Blk) As Dictionary
Set EvlBlkAyPrm = DicByLy(AySelPfx(BlkAySelLy(A, eBlkPrm), "%"))
End Function

Function EvlBlkAySq$(A() As Blk, Prm As Dictionary, Sw As Dictionary)

End Function

Function EvlBlkAySw(A() As Blk, Prm As Dictionary) As Dictionary
Set EvlBlkAySw = EvlSw(AySelPfx(BlkAySelLy(A, eBlkSw), "?"), Prm)
End Function

Function EvlSq$(SqLy$(), Prm As Dictionary, Sw As Dictionary)
Dim Ay$()
    Ay = RsoiLy(SqLy, Sw)
Dim Expr As New Dictionary
    Dim L
    For Each L In AyRmvEmpty(SqLy)
        If IsPfx(L, "Expr") Then
            L = RmvFstTerm(L)
            With Brk(L, " ")
                Expr.Add .S1, .S2
            End With
        Else
            Push Ay, L
        End If
    Next

Dim O$()
    For Each L In Ay
        PushNonEmpty O, EvlSq__ForEachLine(L, Prm, Expr)
    Next
EvlSq = RplVbar(JnCrLf(O))
End Function

Function EvlSq__ForEachLine$(SqLin, Prm As Dictionary, Expr As Dictionary)
Dim L$
    L = SqLin
Dim Sqpi$
    Sqpi = ParseTerm(L)
Dim Pfx$
    Select Case Sqpi
    Case "|Sel":    Pfx = "|Select"
    Case "|SelDis": Pfx = "|Select Distinct"
    Case "Upd":    Pfx = "Update "
    Case "Set":    Pfx = "|  Set"
    Case "Fm":     Pfx = "|  From "
    Case "Left":   Pfx = "|  Left Join "
    Case "Jn":     Pfx = "|  Join "
    Case _
        "WhExpr", _
        "WhInStrLis", _
        "WhInNbrLis", _
        "WhBetStr", _
        "WhBetNbr":
                   Pfx = "|  Where "
    Case "Gp":     Pfx = "|  Group By"
    Case Else
        Stop
    End Select
Dim Rst$
    Select Case Sqpi
    Case _
        "Upd", _
        "Fm", _
        "Left", _
        "Jn", _
        "WhExpr"
                        Rst = L
    Case _
        "|Sel", _
        "|SelDis"
                        Rst = SqtcSel(L, Expr)    ' Sqtc = Sql-Phrase-Itm-Context
    Case "WhInStrLis":  Rst = SqtcWhInStrLis(L, Expr)
    Case "WhInNbrLis":  Rst = SqtcWhInNbrLis(L, Expr)
    Case "WhBetStr":    Rst = SqtcWhBetStr(L, Prm)
    Case "WhBetNbr":    Rst = SqtcWhBetNbr(L, Prm)
    Case "Gp":          Rst = SqtcGp(L, Expr)
    Case Else
        Stop
    End Select
EvlSq__ForEachLine = Pfx & Rst
End Function

Function EvlSw(SwLy$(), Prm As Dictionary) As Dictionary
Dim O As New Dictionary
    O.RemoveAll
    Dim W$()
    Dim SomLinEvaluated As Boolean
    Dim U%
        W = AyRmvEmpty(SwLy)
        U = UB(W)
        SomLinEvaluated = True
    While SomLinEvaluated
        SomLinEvaluated = False
        Dim L
        For Each L In W
            Dim K$
                K = ParseTerm(L)
            With EvlSw__ForEachLine(K, L, Prm, O)
                If .Som Then
                    SomLinEvaluated = True
                    O.Add K, .Bool         '<==
                End If
            End With
        Next
    Wend
If O.Count <> Sz(W) Then Stop
Set EvlSw = O
End Function

Function EvlSw__ForEachLine(K$, OpLin, Prm As Dictionary, Sw As Dictionary) As BoolOpt
'EvlSw__ForEachLine = Evaluate-Line-Switch
'OpLin is up to [Op] not yet parse
If Sw.Exists(K) Then Exit Function
Dim L$
    L = OpLin

Dim O As BoolOpt
    Dim Op$, TermAy$()
        Op = ParseTerm(L)
        TermAy = SplitLvs(L)
    Select Case Op
    Case "OR":  O = EvlSw_OR(TermAy, Prm, Sw)
    Case "AND": O = EvlSw_AND(TermAy, Prm, Sw)
    Case "NE":  O = EvlSw_NE(TermAy, Prm, Sw)
    Case "EQ":  O = EvlSw_EQ(TermAy, Prm, Sw)
    Case Else: Stop
    End Select
EvlSw__ForEachLine = O
End Function

Function EvlSw_AND(TermAy$(), Prm As Dictionary, Sw As Dictionary) As BoolOpt
Dim BoolAy() As Boolean
    Dim I
    For Each I In TermAy
        With EvlSw_Term(I, Prm, Sw)
            If Not .Som Then Exit Function
            Push BoolAy, CBool(.V)
        End With
    Next
EvlSw_AND = SomBool(BoolAy_And(BoolAy))
End Function

Function EvlSw_EQ(TermAy$(), Prm As Dictionary, Sw As Dictionary) As BoolOpt
If Sz(TermAy) <> 2 Then Stop
With EvlSw_T1T2(TermAy(0), TermAy(1), Prm, Sw)
    If Not .Som Then Exit Function
    With .S1S2
        EvlSw_EQ = SomBool(.S1 = .S2)
    End With
End With
End Function

Function EvlSw_NE(TermAy$(), Prm As Dictionary, Sw As Dictionary) As BoolOpt
If Sz(TermAy) <> 2 Then Stop
With EvlSw_T1T2(TermAy(0), TermAy(1), Prm, Sw)
    If Not .Som Then Exit Function
    With .S1S2
        EvlSw_NE = SomBool(.S1 <> .S2)
    End With
End With
End Function

Function EvlSw_OR(TermAy$(), Prm As Dictionary, Sw As Dictionary) As BoolOpt
Dim BoolAy() As Boolean
    Dim I
    For Each I In TermAy
        With EvlSw_Term(I, Prm, Sw)
            If Not .Som Then Exit Function
            Push BoolAy, CBool(.V)
        End With
    Next
EvlSw_OR = SomBool(BoolAy_Or(BoolAy))
End Function

Function EvlSw_T1T2(T1$, T2$, Prm As Dictionary, Sw As Dictionary) As S1S2Opt
Dim S1$, S2$
With EvlSw_Term(T1, Prm, Sw)
    If Not .Som Then Exit Function
    S1 = .V
End With
With EvlSw_Term(T2, Prm, Sw)
    If Not .Som Then Exit Function
    S2 = .V
End With
EvlSw_T1T2 = SomS1S2(S1, S2)
End Function

Function EvlSw_Term(T, Prm As Dictionary, Sw As Dictionary) As VarOpt
'T is for switch-term, which is begin with % or ? or it is *Blank.  % is for parameter & ? is for switch
'  If %, it will evaluated to str
'        if not exist in {Prm}, stop
'  If ?, it will evaluated to bool
'        if not exist in {Switch}, return None
'  Otherwise, just return SomVar(T)
Dim O As VarOpt
    Select Case FstChr(T)
    Case "?"
        If Not Sw.Exists(T) Then Exit Function
        O = SomVar(Sw(T))
    Case "%"
        If Not Prm.Exists(T) Then Stop
        O = SomVar(Prm(T))
    Case "*"
        If T <> "*Blank" Then Stop
        O = SomVar("")
    Case Else
        O = SomVar(T)
    End Select
EvlSw_Term = O
End Function

Function ExprDic(SqLy$()) As Dictionary
Set ExprDic = New Dictionary
Stop
End Function

Function IsRmkLin(I) As Boolean
Dim L$: L = Trim(I)
IsRmkLin = True
If L = "" Then Exit Function
If IsPfx(L, "--") Then Exit Function
IsRmkLin = False
End Function

Function RsoiLin$(SqLin, Sw As Dictionary)
Dim K$, Rst$, Fst$
    With Brk(SqLin, " ")
        K = RmvPfx(.S1, "?")
        Rst = .S2
    End With
    Fst = Left(SqLin, InStr(SqLin, Rst) - 1)
    If IsPfx(Fst, "?") Then Fst = RmvFstChr(Fst) & " "
    If K <> Trim(Fst) Then Stop
Dim O$
    Select Case K
    Case "|Sel", "Gp", "Set":           O = Fst & RsoiSel_or_Gp_or_Set(Rst, Sw)
    Case "AndInStrLis", "AndInNbrLis": O = RsoiAndInXXXLis(Fst, Rst, Sw)
    Case Else: O = SqLin
    End Select
RsoiLin = RplVbar(O)
End Function

Function RsoiLy(SqLy$(), Sw As Dictionary) As String()
'Rsoi:=Remove-Sql-Option-Item
Dim O$()
    Dim L
    For Each L In SqLy
        If IsRmkLin(L) Then
            Push O, L
        Else
            PushNonEmpty O, RsoiLin(L, Sw)
        End If
    Next
RsoiLy = O
End Function

Function SqtcGp$(TermLvs$, Expr As Dictionary)
Dim Ay$()
    Ay = SplitLvs(TermLvs)
Dim ExprAy()
    ExprAy = DicVy(Expr, Ay)
SqtcGp = SqpcGp(ExprAy)
End Function

Function SqtcSel$(TermLvs$, Expr As Dictionary)
'Sqtc:=[S]ql-[T]em[p]late-[C]ontext-For
Dim Fny$()
Dim ExprAy()
    Fny = SplitLvs(TermLvs)
    ExprAy = DicVy(Expr, Fny)
SqtcSel = SqpcSel(Fny, ExprAy)
End Function

Function SqtcWhBetNbr$(TermLvs$, Expr As Dictionary)
SqtcWhBetNbr = SqtcWhBetXXX(TermLvs, Expr, IsStr:=False)
End Function

Function SqtcWhBetStr$(TermLvs$, Expr As Dictionary)
SqtcWhBetStr = SqtcWhBetXXX(TermLvs, Expr, IsStr:=True)
End Function

Function SqtcWhBetXXX$(TermLvs$, Prm As Dictionary, Optional IsStr As Boolean)
Dim T1$, T2$
    With Brk(TermLvs, " ")
        T1 = .S1
        T2 = .S2
    End With
If Not Prm.Exists(T1) Then Stop
If Not Prm.Exists(T2) Then Stop
Dim A1$, A2$
    A1 = Prm(T1)
    A2 = Prm(T2)
Dim C$
    Const C1$ = "Between '?' and '?'"
    Const C2$ = "Between ? and ?"
    C = IIf(IsStr, C1, C2)
SqtcWhBetXXX = FmtQQ(C, A1, A2)
End Function

Function SqtcWhInNbrLis$(TermLvs$, Expr As Dictionary)
SqtcWhInNbrLis = SqtcWhInXXXLis$(TermLvs, Expr, IsStr:=False)
End Function

Function SqtcWhInStrLis$(TermLvs$, Expr As Dictionary)
SqtcWhInStrLis = SqtcWhInXXXLis$(TermLvs, Expr, IsStr:=True)
End Function

Function SqtcWhInXXXLis$(TermLvs$, Expr As Dictionary, IsStr As Boolean)
Dim T1$, T2$
    With Brk(TermLvs, " ")
        T1 = .S1
        T2 = .S2
    End With
If Not Expr.Exists(T1) Then Stop
Dim L$
    Dim Ay$()
    Ay = SplitLvs(T2)
    If IsStr Then Ay = AyQuote(Ay, "'")
    L = JnComma(Ay)
SqtcWhInXXXLis = Expr(T1) & " in (" & L & ")"
End Function

Function Tp$()
'Tp:=Template-Lines
Tp = Join(Array(TpPrm, TpSw, TpDrp, TpT, TpO), Tp_Sep)
End Function

Function Tp_Sep$()
Tp_Sep = vbCrLf & "================================================================"
End Function

Function TpDrp$()
TpDrp = vbCrLf & "Drp Tx TxMbr Crd Div Sto Oup Cnt MbrWs"
End Function

Function TpO$()
'Tpo:=Template-Lines-Output
TpO = Join(Array(TpOOup, TpOCnt, TpOMbrWs), Tp_Sep)
End Function

Function TpOCnt$()
Const L$ = _
"|Sel  Cnt Qty Amt" & _
"|Into #Cnt" & _
"|Fm   #Tx"
TpOCnt = RplVbar(L)
End Function

Function TpOMbrWs$()
Const L$ = _
"|?Sel  Mbr ?Nm ?CNm ?Email ?Phone ?Adr Reg Dist" & _
"|Into #MbrWs" & _
"|Fm   #TxMbr x" & _
"|Left JCMMember a on x.Mbr = a.JCMMCode"
TpOMbrWs = RplVbar(L)
End Function

Function TpOMbrWsOpt$()
Const L$ = ""
TpOMbrWsOpt = RplVbar(L)
End Function

Function TpOOup$()
Const L$ = _
"|Sel  Crd ?Mbr ?Sto ?Div ?DivNm ?Sto ?StoNm ?StoCNm Reg Dist"
TpOOup = RplVbar(L)
End Function

Function TpPrm$()
Const A$ = _
"|%LisDiv     01 02 03  ---" & _
"|%LisCrd     1 2 3 4" & _
"|%LisSto     001 002 003 004" & _
"|%?BrkDiv    1" & _
"|%?BrkSto    1" & _
"|%?BrkCrd    1" & _
"|%?BrkMbr    1" & _
"|%?InclNm    1" & _
"|%?InclAdr   1" & _
"|%?InclPhone 1" & _
"|%?InclEmail 1" & _
"|%SumLvl     W" & _
"|%Fm         20170101" & _
"|%To         20170131"
TpPrm = RplVbar(A)
End Function

Function TpSw$()
Const A$ = _
"|?SelMbr OR  %?InclNm %?InclAdr %?InclPhone %?InclEmail" & _
"|?SelDiv NE  %LisDiv *Blank" & _
"|?SelCrd NE  %LisCrd *Blank" & _
"|?SelSto NE  %LisSto *Blank" & _
"|?LvlY   EQ  %SumLvl Y" & _
"|?LvlM   EQ  %SumLvl M" & _
"|?LvlW   EQ  %SumLvl W" & _
"|?LvlD   EQ  %SumLvl D" & _
"|?InclY  OR  ?LvlD ?LvlW ?LvlM ?LvlY" & _
"|?InclM  OR  ?LvlD ?LvlW ?LvlM" & _
"|?InclW  OR  ?LvlD ?LvlW" & _
"|?InclD  OR  ?LvlD" & _
"|?BrkDte OR  ?LvlY ?LvlM ?LvlW ?LvlD" & _
"|?Mbr    AND %?BrkMbr"
TpSw = RplVbar(A)
End Function

Function TpT$()
'Tpt:=Template-Lines-Temprary
TpT = Join(Array(TpTTx, TpTUpdTx, TpTTxMbr, TpTDiv, TpTSto, TpTCrd), Tp_Sep)
End Function

Function TpTCrd$()
Const L$ = _
"|Sel      Crd CrdNm" & _
"|Into     #Crd" & _
"|Fm       SRCrdTy()" & _
"|?WhInLis Crd ?InCrd"
TpTCrd = RplVbar(L)
End Function

Function TpTDiv$()
Const L$ = _
"|?Sel        Div DivNm" & _
"|Into        #Div" & _
"|Fm          Division" & _
"|?WhInStrLis Div ?InDiv"
TpTDiv = RplVbar(L)
End Function

Function TpTSto$()
Const L$ = _
"|?Sel        Sto StoNm StoCNm" & _
"|Into        #Sto" & _
"|Fm          Location" & _
"|?WhInStoLis Sto ?InSto"
TpTSto = RplVbar(L)
End Function

Function TpTTx$()
Const L$ = _
"|Sel          Crd ?Mbr ?Sto ?Div ?Y ?M ?W ?WD ?WD1 ?D ?Dte Amt Qty Cnt" & _
"|Into         #Tx" & _
"|Fm           SaleHistory" & _
"|WhBetStr     PFm PFm" & _
"|?AndInStrLis Div ?InDiv" & _
"|?AndInNbrLis Crd ?InCrd" & _
"|?AndInStrLis Sto ?InSto" & _
"|Gp           Crd ?Mbr ?Sto ?Div ?Y ?M ?W ?D ?Dte"
TpTTx = RplVbar(L)
End Function

Function TpTTxMbr$()
Const L$ = _
"|SelDis Mbr" & _
"|Into  #TxMbr" & _
"|Fm    #Tx"
TpTTxMbr = RplVbar(L)
End Function

Function TpTUpdTx$()
Const L$ = _
"|Upd  #Tx" & _
"|Set ?WD"
TpTUpdTx = RplVbar(L)
End Function

Sub XDmpBlkI(I%)
BlkDmp YBlkI(I)
End Sub

Sub XDmpBlkTy()
AyDmp AyMapIntoSy(YBlkTyAy, "BlkTyToStr")
End Sub

Sub XDmpBlkTyI(I%)
Debug.Print BlkTyToStr(YBlkI(I).Ty)
End Sub

Sub XDmpChkBlkAy()
StrOptDmp YChkBlkAy
End Sub

Sub XDmpChkPrm()
ChkRsltDmp YChkPrm
End Sub

Sub XDmpChkPrmLy()
AyDmp YChkPrmLy
End Sub

Sub XDmpChkSw()
ChkRsltDmp YChkSw
End Sub

Sub XDmpChkSw__ForEachLine()
ChkRsltDmp YChkSw__ForeachLine
End Sub

Sub XDmpChkSwLy()
AyDmp YChkSwLy
End Sub

Sub XDmpClyPrm1()
ChkRsltDmp YClyPrm1
End Sub

Sub XDmpClyPrm2()
ChkRsltDmp YClyPrm2
End Sub

Sub XDmpCrAy(A() As ChkRslt)
Dim J%
For J = 0 To ChkRsltUB(A)
Next
End Sub

Sub XDmpDicPrm()
DicDmp YDicPrm
End Sub

Sub XDmpDicSw()
DicDmp YDicSw
End Sub

Sub XDmpEvlBlkAyPrm()
DicDmp YEvlBlkAyPrm
End Sub

Sub XDmpEvlBlkAySw()
'DicDmp YEvlBlkAySw
'DicDmp YDicSw
Dim O$()
PushAy O, DrsLy(DicJn(Array(YEvlBlkAySw, YDicSw)))
PushAy O, DicLy(YDicPrm)
AyBrw O
End Sub

Sub XDmpEvlSqI(I%)
Debug.Print YEvlSqI(I)
End Sub

Sub XDmpEvlSw()
Dim O$()
PushAy O, DrsLy(DicJn(Array(YEvlSw, YDicSw)))
PushAy O, DicLy(YDicPrm)
AyBrw O
End Sub

Function YBlkAy() As Blk()
YBlkAy = BlkBrk(Tp)
End Function

Function YBlkI(I%) As Blk
Dim A() As Blk: A = YBlkAy
YBlkI = A(I)
End Function

Function YBlkTyAy() As eBlkTy()
Dim J%, A() As Blk
A = YBlkAy
Dim O() As eBlkTy
For J = 0 To BlkUB(A)
    Push O, A(J).Ty
Next
YBlkTyAy = O
End Function

Function YBlkTyI(I%) As eBlkTy
YBlkTyI = YBlkI(I).Ty
End Function

Function YChkBlkAy() As StrOpt
YChkBlkAy = ChkBlkAy(YBlkAy)
End Function

Function YChkPrm() As ChkRslt()
YChkPrm = ChkPrm(YLyPrm)
End Function

Function YChkPrmLy() As String()
YChkPrmLy = ChkRsltPut(YLyPrm, YChkPrm)
End Function

Function YChkSw() As ChkRslt()
YChkSw = ChkSw(YLySw, YDicPrm, YDicSw)
End Function

Function YChkSw__ForDupKey() As ChkRslt()
YChkSw__ForDupKey = ChkSw__ForDupKey(YLySw)
End Function

Function YChkSw__ForeachLine() As ChkRslt()
YChkSw__ForeachLine = ChkSw__ForEachLine(YLySw, YDicPrm, YDicSw)
End Function

Function YChkSwLy() As String()
YChkSwLy = ChkRsltPut(YLySw, YChkSw)
End Function

Function YClyPrm1() As ChkRslt()
YClyPrm1 = ChkPrm__ForEachLine(YLyPrm)
End Function

Function YClyPrm2() As ChkRslt()
YClyPrm2 = ChkPrm__ForDupKey(YLyPrm)
End Function

Function YDicPrm() As Dictionary
Set YDicPrm = DicByLy(YLyPrm, IgnoreDup:=True)
End Function

Function YDicSw() As Dictionary
Set YDicSw = DicByLy(YLySw, IgnoreDup:=True)
End Function

Function YEvlBlkAyPrm() As Dictionary
Set YEvlBlkAyPrm = EvlBlkAyPrm(YBlkAy)
End Function

Function YEvlBlkAySw() As Dictionary
Set YEvlBlkAySw = EvlBlkAySw(YBlkAy, YEvlBlkAyPrm)
End Function

Function YEvlSqI$(I%)

YEvlSqI = EvlSq(YLySqI(I), YDicPrm, YEvlSw)
End Function

Function YEvlSw() As Dictionary
Set YEvlSw = EvlSw(YLySw, YDicPrm)
End Function

Function YLyPrm() As String()
YLyPrm = BlkAySelLy(YBlkAy, eBlkPrm)
End Function

Function YLySqAy() As Variant()
YLySqAy = BlkAySelLyAy(YBlkAy, eBlkSq)
End Function

Function YLySqI(I%) As String()
Dim A() As Variant: A = YLySqAy
YLySqI = A(I)
End Function

Function YLySw() As String()
YLySw = BlkAySelLy(YBlkAy, eBlkSw)
End Function

Function YNBlk%()
YNBlk = BlkSz(YBlkAy)
End Function

Function YNSq%()
YNSq = Sz(YLySqAy)
End Function

Function YUBlk%()
YUBlk = BlkUB(YBlkAy)
End Function

Function YUSq%()
YUSq = UB(YLySqAy)
End Function

Private Function RsoiAndInXXXLis$(Fst$, Rst$, Sw As Dictionary)
Dim S2$, S1$
    With Brk(Rst, " ")
        S1 = .S1
        S2 = .S2
    End With
If Not IsPfx(S2, "?") Then RsoiAndInXXXLis = Fst & Rst
If Not Sw.Exists(S2) Then Stop
If Sw(S2) Then RsoiAndInXXXLis = Fst & S1 & " " & RmvPfx(S2, "?")
End Function

Private Function RsoiSel_or_Gp_or_Set$(Rst$, Sw As Dictionary)
Const CSub$ = "RsoiSel_or_Gp_or_Set"
Dim O$()
    Dim I
    For Each I In SplitLvs(Rst)
        If IsPfx(I, "?") Then
            If Not Sw.Exists(I) Then Er CSub, "Sw Dic", FmtQQ("Key[?] missing", I)
            If Sw(I) Then Push O, RmvFstChr(I)
        Else
            Push O, I
        End If
    Next
RsoiSel_or_Gp_or_Set = JnSpc(O)
End Function

Function ChkSq__Tst()
Dim Prm As New Dictionary
Dim Expr As New Dictionary
Dim Sw As New Dictionary
Dim SqLy$()
Dim Act() As ChkRslt
    With Sw
        .RemoveAll
        .Add "?Mbr", True
        .Add "?Sto", True
        .Add "?Crd", True
        .Add "?Div", False
        .Add "?Y", False
        .Add "?M", True
        .Add "?W", True
        .Add "?D", True
        .Add "?Dte", False
        .Add "?WD", False
        .Add "?WD1", False
        .Add "?InDiv", True
        .Add "?InSto", True
        .Add "?InCrd", True
    End With
    With Prm
        .RemoveAll
    End With
    With Expr
        .RemoveAll
    End With
    SqLy = SplitVBar( _
        "|Sel          Crd ?Mbr ?Sto ?Div ?Y ?M ?W ?WD ?WD1 ?D ?Dte Amt Qty Cnt" & _
        "|Into         #Tx" & _
        "|Fm           SaleHistory" & _
        "|WhBetStr     %Fm %To1" & _
        "|?AndInStrLis Div ?InDiv" & _
        "|?AndInNbrLis Crd ?InCrd" & _
        "|?AndInStrLis Sto ?InSto" & _
        "|Gp           Crd ?Mbr ?Sto ?Div ?Y ?M ?W ?D ?Dte")
    Act = ChkSq(SqLy, Sw, Prm)
    ChkRsltDmp Act
    AyDmp ChkRsltPut(SqLy, Act)
Stop
End Function

Sub ChkSw__Tst()
Dim Act() As ChkRslt
Dim Exp() As ChkRslt
Dim Prm As Dictionary
Dim Sw As Dictionary
    Stop
Dim SwBlk$
    Set Prm = DicByLines(TpPrm)
    Act = SqTp.ChkSw(SplitCrLf(TpSw), Prm, Sw)
'    If Not StrOptIsEq(Act, Exp) Then StrBrw Act.Str

End Sub

Sub EvlSq__Tst()
Dim SqlBlk$, Prm As Dictionary, Expr As Dictionary
Dim Exp$, Act$
'---
SqlBlk = ""

End Sub

Sub EvlSw__Tst()
Dim SwLy$()
Dim Prm As Dictionary
Dim Act As Dictionary
Dim Exp As New Dictionary

    Set Prm = DicByLines(RplVbar( _
        "%BrkCrd     False" & _
        "|%BrkDiv     False" & _
        "|%BrkMbr     False" & _
        "|%BrkSto     False" & _
        "|%CrdLis    " & _
        "|%StoLis    " & _
        "|%DivLis     01 02" & _
        "|%FmDte      20170101" & _
        "|%ToDte      20170131" & _
        "|%SumLvl     M" & _
        "|%InclAdr    False" & _
        "|%InclNm     False" & _
        "|%InclPhone  False" & _
        "|%InclEmail  False" _
        ))
    
    SwLy = SplitCrLf(RplVbar( _
        "|?InDiv    NE  %DivLis *Blank" & _
        "|?InCrd    NE  %CrdLis *Blank" & _
        "|?InSto    NE  %StoLis *Blank" & _
        "|?LvlY     EQ  %SumLvl Y" & _
        "|?LvlM     EQ  %SumLvl M" & _
        "|?LvlW     EQ  %SumLvl W" & _
        "|?LvlD     EQ  %SumLvl D" & _
        "|?Y        OR  ?LvlD ?LvlW ?LvlM ?LvlY" & _
        "|?M        OR  ?LvlD ?LvlW ?LvlM" & _
        "|?W        OR  ?LvlD ?LvlW" & _
        "|?WD       OR  ?W" & _
        "|?WD1      OR  ?W" & _
        "|?D        OR  ?LvlD" & _
        "|?Dte      OR  ?LvlD" & _
        "|?Sto      OR  %BrkMbr" & _
        "|?Crd      OR  %BrkCrd" & _
        "|?Mbr      OR  %BrkMbr" & _
        "|?MbrWs    OR  %InclNm %InclAdr %InclEmail %InclPhone" & _
        "|?SqlMbrWs AND %BrkMbr ?MbrWs" _
        ))
Set Act = EvlSw(SwLy, Prm)
With Exp
    .RemoveAll
    .Add "XX", True
End With
DicAssertIsEq Act, Exp
End Sub

Sub RsoiLy__Tst()
Dim Sw As New Dictionary
Dim SqLy$()
    With Sw
'        .Add "?Mbr", True
        .Add "?Sto", True
        .Add "?Crd", True
        .Add "?Div", False
        .Add "?Y", False
        .Add "?M", True
        .Add "?W", True
        .Add "?D", True
        .Add "?Dte", False
        .Add "?WD", False
        .Add "?WD1", False
        .Add "?InDiv", True
        .Add "?InSto", True
        .Add "?InCrd", True
    End With
    SqLy = SplitVBar( _
        "|Sel          Crd ?Mbr ?Sto ?Div ?Y ?M ?W ?WD ?WD1 ?D ?Dte Amt Qty Cnt" & _
        "|Into         #Tx" & _
        "|Fm           SaleHistory" & _
        "|WhBetStr     PFm PFm" & _
        "|?AndInStrLis Div ?InDiv" & _
        "|?AndInNbrLis Crd ?InCrd" & _
        "|?AndInStrLis Sto ?InSto" & _
        "|Gp           Crd ?Mbr ?Sto ?Div ?Y ?M ?W ?D ?Dte")

AyDmp RsoiLy(SqLy, Sw)
End Sub

