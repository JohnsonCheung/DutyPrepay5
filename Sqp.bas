Attribute VB_Name = "Sqp"
Option Compare Database
Option Explicit

Function FmtSql$(SqlTp$, Dic As Dictionary)

End Function

Property Get MulSqlDrp$(TblNmLvs$)

End Property

Function SqpAnd$(Expr$)
If Expr = "" Then Exit Function
SqpAnd = "|    And " & Expr
End Function

Function SqpcGp$(ExprVblAy)
'Sqpc = Sql-Phrase-Context
SqpcGp = JnComma(AyAddPfx(VblAyAlignL(ExprVblAy, 4, 6), "|"))
End Function

Function SqpcSel$(Fny, ExprVblAy)
'Sqpc = Sql-Phrase-Context
Assert_Fny_ExprVblAy Fny, ExprVblAy
Dim AlignedExprAy$()
Dim AlignedFny$()
    AlignedExprAy = VblAyAlignL(ExprVblAy, 4, 6)
    AlignedFny = AyAlignL(Fny)
Dim mS1S2Ay() As S1S2
    mS1S2Ay = S1S2Ay(AlignedExprAy, AlignedFny)
Dim O$()
    O = S1S2AyConcat(mS1S2Ay, " ")
    O = AyAddPfx(O, "|")
SqpcSel = Join(O, ",")
End Function

Function SqpExprIn$(Expr$, InLis$)
If InLis = "" Then Exit Function
SqpExprIn = FmtQQ("? in (?)", Expr, InLis)
End Function

Function SqpFm$(T)
SqpFm = "|  From " & T
End Function

Function SqpGp$(ExprVblAy)
VblAyAssertIsVdt ExprVblAy
SqpGp = "|  Group By" & SqpcGp(ExprVblAy)
End Function

Function SqpInto$(T)
SqpInto = "|  Into " & T
End Function

Function SqpSel$(Fny, ExprVblAy)
SqpSel = "Select" & SqpcSel(Fny, ExprVblAy)
End Function

Function SqpSelDis$(Fny, ExprVblAy)
SqpSelDis = "Select Distinct|" & SqpcSel$(Fny, ExprVblAy)
End Function

Function SqpSelDisFldLvs$(FldLvs, ExprVblAy)
Dim Fny$(): Fny = SplitLvs(FldLvs)
SqpSelDisFldLvs = SqpSelDis(Fny, ExprVblAy)
End Function

Function SqpSelFldLvs$(FldLvs, ExprVblAy)
Dim Fny$(): Fny = SplitLvs(FldLvs)
SqpSelFldLvs = SqpSel(Fny, ExprVblAy)
End Function

Function SqpSet$(FldLvs$, ExprVblAy)
Dim Fny$(): Fny = SplitLvs(FldLvs)
VblAyAssertIsVdt ExprVblAy
If Sz(Fny) <> Sz(ExprVblAy) Then Stop
Dim AFny$()
Dim AExprAy$()
    Dim W%
    AFny = AyAddPfx(AyAlignL(Fny), "    ")
    W = Len(AFny(0)) + 3
    AExprAy = VblAyAlignL(0, W)
Dim O$()
    Dim J%
    For J = 0 To UB(AFny)
        Push O, "|" & AFny(J) & " = " & AExprAy(J)
    Next
SqpSet = "|  Set" & JnComma(O)
End Function

Function SqpUpd$(T)
SqpUpd = "Update " & T
End Function

Function SqpWh$(Expr)
SqpWh = "|  Where " & Expr
End Function

Function SqpWhBetStr$(FldNm$, FmStr$, ToStr$)
SqpWhBetStr = FmtQQ("|  Where ? Between '?' and '?'", FldNm, FmStr, ToStr)
End Function

Private Sub Assert_Fny_ExprVblAy(Fny, ExprVblAy)
Const CSub$ = "Assert_Fny_ExprVblAy"
AssertIsAy Fny
VblAyAssertIsVdt ExprVblAy
If UB(Fny) <> UB(ExprVblAy) Then Er CSub, "?: UB-{Fny} <> UB-{ExprAy}", CSub, UB(Fny), UB(ExprVblAy)
AyAssertNoEmptyEle Fny
AyAssertNoEmptyEle ExprVblAy
End Sub

Private Function ZCpy()

'===================================================
'Option Compare Database
'Option Explicit
'Private ExprVblAy
'Private Fny
'Property Get MulSqlDrp$(TblNmLvs$)
'
'End Property
'Function SqpGp$(ExprVblAy)
'
'End Function
'Function SqpInto$(T)
'SqpInto = "|  Into " & T
'End Function
'Function SqpSel$(Fny_, ExprVblAy_)
'AssertIsAy Fny_
'AssertIsAy ExprVblAy_
'Fny = Fuy_
'ExprVblAy = ExprVblAy_
'If UB(Fny) <> UB(ExprVblAy) Then Stop
'Dim B() As S1S2
'    Dim B1() As S1S2
'    Dim B2() As S1S2
'    Dim B3() As S1S2
'    Dim B4() As S1S2
'    B1 = S1S2Ay(Fny, ExprVblAy)
'    B2 = Z1__SetFldNmPfxIsDot(B1)
'    B3 = Z2__RmvEmptyExprItm(B2)
'    B4 = Z3__RmvFldNmPfxIsQuestionMrk(B3)
'     B = Z4__RmvTermWithDot(B4)
'Dim F1$()
'Dim E1$()
'    F1 = S1S2AyS1Ay(B)
'    E1 = S1S2AyS2Ay(B)
'Dim F2$()
'Dim E2$()
'    F2 = AyAlignL(F1)
'    E2 = Z5__VblAyAlignL(E1)
'Dim E3$()
'    E3 = Z6__ExprLinesAyTab(E2, 4)
'
'Dim O$()
'    O = S1S2AyConcat(S1S2Ay(E3, F2), " ")
'SqpSel = "Select|" & Join(O, ",|")
'End Function
'
'Private Function Z1__SetFldNmPfxIsDot(A() As S1S2) As S1S2()
'Dim J%, O() As S1S2
'O = A
'For J = 0 To S1S2UB(O)
'    If FstChr(A(J).S1) = "." Then
'        O(J).S1 = RmvFstChr(O(J).S1)
'        O(J).S2 = O(J).S1
'    End If
'Next
'Z1__SetFldNmPfxIsDot = O
'End Function
'
'Private Function Z2__RmvEmptyExprItm(A() As S1S2) As S1S2()
'Dim O() As S1S2
'Dim F$(), E$()
'Dim J%
'For J = 0 To S1S2UB(A)
'    If A(J).S2 <> "" Then
'        S1S2Push O, A(J)
'    End If
'Next
'Z2__RmvEmptyExprItm = O
'End Function
'
'Private Function Z3__RmvFldNmPfxIsQuestionMrk(A() As S1S2) As S1S2()
'Dim O() As S1S2
'O = A
'Dim J%
'For J = 0 To S1S2UB(O)
'    With O(J)
'        If A(J).S2 <> "" Then
'            If IsPfx(.S1, "?") Then
'                .S1 = RmvFstChr(0.1)
'            End If
'        End If
'    End With
'Next
'Z3__RmvFldNmPfxIsQuestionMrk = O
'End Function
'
'Private Function Z4__RmvTermWithDot(A() As S1S2) As S1S2()
'Dim J%, O() As S1S2
'O = A
'For J = 0 To S1S2UB(O)
'    If HasSubStr(O(J).S1, ".") Then
'        O(J).S1 = TakAftRev(O(J).S1, ".")
'        O(J).S1 = RmvPfx(O(J).S1, "?")
'    End If
'Next
'Z4__RmvTermWithDot = O
End Function

Private Function ZZExprVblAy() As String()
ZZExprVblAy = Sy("F1-Expr", "F2-Expr   AA|BB    X|DD       Y", "F3-Expr  x")
End Function

Private Function ZZFny() As String()
ZZFny = SplitSpc("F1 F2 F3xxxxx")
End Function

Private Sub FmtSql__Tst()
Dim Tp$: Tp = "Select" & _
"|{?Sel}" & _
"|    {ECrd} Crd," & _
"|    {EAmt} Amt," & _
"|    {EQty} Qty," & _
"|    {ECnt} Cnt," & _
"|  Into #Tx" & _
"|  From SaleHistory" & _
"|  Where SHDate Between '{PFm}' and '{PTo}'" & _
"|    And {EDiv} in ({InDiv})" & _
"|  Group By" & _
"|{?Gp}" & _
"|?M   {ETxM}," & _
"|?W   {ETxW}," & _
"|?D   {ETxD}"
'SR_ = Sales Report
Const ETxWD$ = _
"CASE WHEN TxWD1 = 1 then 'Sun'" & _
"|ELSE WHEN TxWD1 = 2 THEN 'Mon'" & _
"|ELSE WHEN TxWD1 = 3 THEN 'Tue'" & _
"|ELSE WHEN TxWD1 = 4 THEN 'Mon'" & _
"|ELSE WHEN TxWD1 = 5 THEN 'Thu'" & _
"|ELSE WHEN TxWD1 = 6 THEN 'Fri'" & _
"|ELSE WHEN TxWD1 = 7 THEN 'Sat'" & _
"|ELSE Null" & _
"|END END END END END END END"
Dim D As New Dictionary
With D
    .Add "ECrd", "Line-1|Line-2"
    .Add "EAmt", "Sum(SHTxDate)"
    
End With
Dim Act$: Act = FmtSql(Tp, D)
Dim Exp$: Exp = ""
Debug.Assert Act = Exp
End Sub

Private Sub SqpSel__Tst()
Debug.Print RplVbar(SqpSel(ZZFny, ZZExprVblAy))
End Sub

