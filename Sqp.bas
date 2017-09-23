Attribute VB_Name = "Sqp"
Option Compare Database
Option Explicit
Private ExprVblAy_
Private Fny_
'

Property Get MulSqlDrp$(TblNmLvs$)

End Property

Function SqpGp$(ExprVblAy)
ExprVblAy_ = ExprVblAy
SqpGp = "|  Group By|" & Join(AlignedExprAy, ",|")
End Function
Function SqpAnd$(Expr$)
If Expr = "" Then Exit Function
SqpAnd = "|    And " & Expr
End Function
Function SqpAndIn$(ExprVblAy, InLisAy)
Dim O$(), J%
O = AyAddPfx(AlignedExprAy, "|    And ")
For J = 0 To UB(O)
    O(J) = O(J) & " in (" & InLisAy(J) & ")"
Next
SqpAndIn = JnComma(O)
End Function
Function SqpWhBetStr$(FldNm$, FmStr$, ToStr$)
SqpWhBetStr = FmtQQ("|  Where ? Between '?' and '?'", FldNm, FmStr, ToStr)
End Function
Function SqpFm$(T)
SqpFm = "|  From " & T
End Function

Function SqpInto$(T)
SqpInto = "|  Into " & T
End Function

Function SqpSel$(Fny, ExprVblAy)
AssertIsAy Fny
AssertIsVblAy ExprVblAy
Fny_ = Fny
ExprVblAy_ = ExprVblAy
If UB(Fny) <> UB(ExprVblAy) Then Stop
Dim O$()
    O = S1S2AyConcat(S1S2Ay(AlignedExprAy(4, 2), AlignedFny), " ")
SqpSel = "Select|" & Join(O, ",|")
End Function
Private Sub SqpSel__Tst()
Debug.Print RplVBar(SqpSel(ZZFny, ZZExprVblAy))
End Sub
Private Sub AAA()
SqpSel__Tst
End Sub

Private Function AlignedExprAy(Optional FstVblNSpc%, Optional RstVblNSpc%) As String()
Dim O$()
Dim ExprLines
For Each ExprLines In ExprLinesAyAlignL(FstVblNSpc, RstVblNSpc)
    Push O, VblIndent(ExprLines, FstVblNSpc, RstVblNSpc)
Next
AlignedExprAy = O
End Function

Private Function AlignedFny() As String()
AlignedFny = AyAlignL(Fny_)
End Function

Private Function ExprLinesAyAlignL(Optional FstVblNSpc%, Optional RstVblNSpc%) As String()
Dim ExprWdt%
    ExprWdt = ExprLinesAyWdt(FstVblNSpc%, RstVblNSpc%)
Dim O$(), W%
Dim I
For Each I In ExprVblAy_
    W = ExprWdt - Len(VblLasLin(I)) + 1
    If W < 1 Then W = 1
    Push O, I & Space(W)
Next
ExprLinesAyAlignL = O
End Function

Private Function ExprLinesAyWdt%(Optional FstVblNSpc%, Optional RstVblNSpc%)
Dim W%, J%, O%, A$()
A = ExprVblAy_
Dim I
For Each I In A
    W = VblWdt(VblIndent(I, FstVblNSpc, RstVblNSpc))
    If W > O Then O = W
Next
ExprLinesAyWdt = O
End Function


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
'    E2 = Z5__ExprLinesAyAlignL(E1)
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

Private Sub ZZSetPrm()
ExprVblAy_ = ZZExprVblAy
Fny_ = ZZFny
End Sub

Private Sub AlignedExprAy__Tst()
ZZSetPrm
AyDmp AlignedExprAy
End Sub
