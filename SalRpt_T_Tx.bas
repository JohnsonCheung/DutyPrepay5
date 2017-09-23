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
'Prm
'    DivLis . 01 02 03
'    CrdLis . 1 2 3 4
'    StoLis . 001 002 003 004
'    ?BrkDiv . 1
'    ?BrkSto . 1
'    ?BrkCrd . 1
'    ?BrkMbr . 0
'    ?InclNm . 1
'    ?InclAdr  . 1
'    ?InclPhone . 1
'    ?InclEmail . 1
'    SumLvl .Y
'    Fm . 20170101
'    To . 20170131
Sub SR_T_Tx_Sql__Tst()
Debug.Print RplVBar(SR_T_Tx_Sql)
End Sub
Private Sub AAA()
SR_T_Tx_Sql__Tst
End Sub
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

Private Function InclSelDiv() As Boolean
InclSelDiv = Prm.DivLis <> ""
End Function
Private Function InclSelSto() As Boolean
InclSelSto = Prm.StoLis <> ""
End Function
Private Function InclSelCrd() As Boolean
InclSelCrd = Prm.CrdLis <> ""
End Function
Private Function ZSelFny() As String()
Dim O$()
Push O, "Crd"
Push O, "Amt"
Push O, "Qty"
Push O, "Cnt"
If InclFldMbr Then Push O, "Mbr"
If InclFldDiv Then Push O, "Div"
If InclFldSto Then Push O, "Sto"
If InclFldTxY Then Push O, "TxY"
If InclFldTxM Then Push O, "TxM"
If InclFldTxD Then Push O, "TxD"
If InclFldTxDte Then Push O, "TxDte"
ZSelFny = O
End Function

Private Function ZSelExpr() As String()
Dim O$()
Push O, ECrd
Push O, EAmt
Push O, EQty
Push O, ECnt
If InclFldMbr Then Push O, EMbr
If InclFldDiv Then Push O, EDiv
If InclFldSto Then Push O, ESto
If InclFldTxY Then Push O, ETxY
If InclFldTxM Then Push O, ETxM
If InclFldTxD Then Push O, ETxD
If InclFldTxDte Then Push O, ETxDte
ZSelExpr = O
End Function

Private Function InclFldMbr() As Boolean
InclFldMbr = Prm.BrkMbr
End Function

Private Function InclFldDiv() As Boolean
InclFldDiv = Prm.BrkDiv
End Function

Private Function InclFldSto() As Boolean
InclFldSto = Prm.BrkDiv
End Function

Private Function InclFldTxY() As Boolean
Select Case Prm.SumLvl
Case "D", "W", "M", "Y": InclFldTxY = True
End Select
End Function

Private Function InclFldTxM() As Boolean
Select Case Prm.SumLvl
Case "D", "W", "M": InclFldTxM = True
End Select
End Function

Private Function InclFldTxW() As Boolean
Select Case Prm.SumLvl
Case "D", "W": InclFldTxW = True
End Select
End Function

Private Function InclFldTxD() As Boolean
Select Case Prm.SumLvl
Case "D": InclFldTxD = True
End Select
End Function

Private Function InclFldTxDte() As Boolean
InclFldTxDte = InclFldTxD
End Function

Private Function ZAnd$()
Dim O$()
If InclSelDiv Then Push O, ZAndDiv
If InclSelSto Then Push O, ZAndSto
If InclSelCrd Then Push O, ZAndCrd
O = AyAddPfx(O, "|    And ")
ZAnd = Join(O)
End Function

Private Function ZAndDiv$()
ZAndDiv = FmtQQ("? in (?)", EDiv, InDiv)
End Function
Private Function ZAndCrd$()
ZAndCrd = FmtQQ("? in (?)", ECrd, InCrd)
End Function
Private Function ZAndSto$()
ZAndSto = FmtQQ("? in (?)", ESto, InSto)
End Function
Private Function InSto$()
Dim A$(): A = SplitSpc(Prm.StoLis)
Dim J%
For J = 0 To UB(A)
    If Len(A(J)) <> 3 Then Stop
Next
InSto = JnComma(AyQuote(A, "'"))
End Function
Private Function InDiv$()
Dim A$(): A = SplitSpc(Prm.DivLis)
Dim J%
For J = 0 To UB(A)
    If Len(A(J)) <> 2 Then Stop
Next
InDiv = JnComma(AyQuote(A, "'"))
End Function
Private Function InCrd$()
Dim A$(): A = SplitSpc(Prm.DivLis)
InCrd = JnComma(A)
End Function
Private Function ZGpExprAy() As String()
Dim O$()
Push O, ECrd
If InclFldMbr Then Push O, EMbr
If InclFldDiv Then Push O, EDiv
If InclFldSto Then Push O, ESto
If InclFldTxY Then Push O, ETxY
If InclFldTxM Then Push O, ETxM
If InclFldTxD Then Push O, ETxD
If InclFldTxDte Then Push O, ETxDte
ZGpExprAy = O
End Function

'Sql.X.T.Tx
'    Sel    # Crd Amt Qty Cnt ?Mbr ?Div ?Sto ?TxY ?TxM ?TxW ?TxD ?TxDte
'    Fm     # SalesHistory
'    Wh     # SHSDate between '{Prm.Fm}' and '{Prm.To}'
'    And    # ?Div ?Sto ?Crd
'    Gp     # Crd ?Mbr ?Div ?Sto ?TxY ?TxM ?TxW ?TxD ?TxDte
'Sql.X.T.Tx.Sel
'    Crd          $ Expr.Crd
'    Amt .Sum(SHAmount)
'    Qty .Sum(SHQty)
'    Cnt .Count(SHInvoice + SHSDate + SHRef)
'    ?Mbr ?BrkMbr . JCMMCode
'    ?Div ?BrkDiv $ Expr.Div
'    ?Sto ?BrkSto $ Expr.Sto
'    ?TxY   ?InclY . SUBSTR(SHSDate,1,4)
'    ?TxM   ?InclM . SUBSTR(SHSDate,5,2)
'    ?TxW   ?InclW . TxW-Expr
'    ?TxD   ?InclD . SUBSTR(SHSDate,7,2)
'    ?TxDte ?InclD . SUBSTR(SHSDate,1,4)+'/'+SUBSTR(SHSDate,5,2)+'/'+SUBSTR(SHSDate,7,2)
'Sql.X.T.Tx.And
'    ?Div ?SelDiv $In Expr.Div Div
'    ?Crd ?SelCrd $In Expr.Crd Crd
'    ?Sto ?SelSto $In Expr.Sto Sto
'Sql.X.T.Tx.And.Print Crd; Crd; !LvsJnComma; Prm.CrdLis
'Sql.X.T.Tx.And.Print Div; Div; !LvsJnQuoteComma; Prm.DivLis
'Sql.X.T.Tx.And.Print Sto; Sto; !LvsJnQuoteComma; Prm.StoLis
