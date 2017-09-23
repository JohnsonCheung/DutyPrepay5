Attribute VB_Name = "SalRpt_T_Tx"
Option Compare Database
Option Explicit
Const ECnt$ = "Count(SHInvoice + SHSDate + SHRef)"
Const EAmt$ = "Sum(SHAmount)"
Const EQty$ = "Sum(SHQty)"
Const EMbr$ = "Mbr-Expr"
Const EDiv$ = "Div-Expr"
Const ESto$ = "Sto-Expr"
Const ETxY$ = "SUBSTR(SHSDate,1,4)"
Const ETxM$ = "SUBSTR(SHSDate,5,2)"
Const ETxW$ = "TxW-Expr"
Const ETxD$ = "SUBSTR(SHSDate,7,2)"
Const ETxDte$ = "SUBSTR(SHSDate,1,4)+'/'+SUBSTR(SHSDate,5,2)+'/'+SUBSTR(SHSDate,7,2)"
Private Prm As SR_Prm

Function ECrd$()
ECrd = CrdExpr(Prm.CrdLis)
End Function

Function SR_T_Tx_Sql$()
Prm = SR_Prm
SR_T_Tx_Sql = _
SqpSel(ZSelFny, ZSelExpr) & _
SqpInto("#Tx") & _
SqpFm("SaleHistory") & _
SqpWhBetStr("SHDate", Prm.FmDte, Prm.ToDte) & _
SqpAnd(ZAndCrd) & _
SqpGp(ZGpExprAy)
End Function

Private Sub AAA()
SR_T_Tx_Sql__Tst
End Sub

Private Function InCrd$()
Dim A$(): A = SplitSpc(Prm.DivLis)
InCrd = JnComma(A)
End Function

Private Function InDiv$()
Dim A$(): A = SplitSpc(Prm.DivLis)
Dim J%
For J = 0 To UB(A)
    If Len(A(J)) <> 2 Then Stop
Next
InDiv = JnComma(AyQuote(A, "'"))
End Function

Private Function InSto$()
Dim A$(): A = SplitSpc(Prm.StoLis)
Dim J%
For J = 0 To UB(A)
    If Len(A(J)) <> 3 Then Stop
Next
InSto = JnComma(AyQuote(A, "'"))
End Function

Private Function ZAnd$()
Dim O$()
If SRS_InclSelDiv Then Push O, ZAndDiv
If SRS_InclSelSto Then Push O, ZAndSto
If SRS_InclSelCrd Then Push O, ZAndCrd
O = AyAddPfx(O, "|    And ")
ZAnd = Join(O)
End Function

Private Function ZAndCrd$()
ZAndCrd = FmtQQ("? in (?)", ECrd, InCrd)
End Function

Private Function ZAndDiv$()
ZAndDiv = FmtQQ("? in (?)", EDiv, InDiv)
End Function

Private Function ZAndSto$()
ZAndSto = FmtQQ("? in (?)", ESto, InSto)
End Function

Private Function ZGpExprAy() As String()
Dim O$()
Push O, ECrd
If SRS_InclFldMbr Then Push O, EMbr
If SRS_InclFldDiv Then Push O, EDiv
If SRS_InclFldSto Then Push O, ESto
If SRS_InclFldTxY Then Push O, ETxY
If SRS_InclFldTxM Then Push O, ETxM
If SRS_InclFldTxD Then Push O, ETxD
If SRS_InclFldTxDte Then Push O, ETxDte
ZGpExprAy = O
End Function

Private Function ZSelExpr() As String()
Dim O$()
Push O, ECrd
Push O, EAmt
Push O, EQty
Push O, ECnt
If SRS_InclFldMbr Then Push O, EMbr
If SRS_InclFldDiv Then Push O, EDiv
If SRS_InclFldSto Then Push O, ESto
If SRS_InclFldTxY Then Push O, ETxY
If SRS_InclFldTxM Then Push O, ETxM
If SRS_InclFldTxD Then Push O, ETxD
If SRS_InclFldTxDte Then Push O, ETxDte
ZSelExpr = O
End Function

Private Function ZSelFny() As String()
Dim O$()
Push O, "Crd"
Push O, "Amt"
Push O, "Qty"
Push O, "Cnt"
If SRS_InclFldMbr Then Push O, "Mbr"
If SRS_InclFldDiv Then Push O, "Div"
If SRS_InclFldSto Then Push O, "Sto"
If SRS_InclFldTxY Then Push O, "TxY"
If SRS_InclFldTxM Then Push O, "TxM"
If SRS_InclFldTxD Then Push O, "TxD"
If SRS_InclFldTxDte Then Push O, "TxDte"
ZSelFny = O
End Function

Sub SR_T_Tx_Sql__Tst()
Debug.Print RplVBar(SR_T_Tx_Sql)
End Sub

