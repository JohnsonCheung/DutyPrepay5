Attribute VB_Name = "Sql3v1"
Option Explicit
Option Compare Database


Private Sub AA_Sql__Flow()

'-- Rmk: -- is remark
'-- 3Lvl: alway 3 level
'-- 4spc: Lvl1 has no space, Lvl2 has exactly 4 space and Lvl3 always have 8 space
'-- NoSpcInNm: Lvl2 (name), cannot have space
'-- Lvl1: is namespace, use do to separate
'-- Lvl2: is name.  That means is always under a namespace
'-- Root Ns: fst non remark line is root ns
'-- Lvl3: is expression
'-- Lvl2Nm-?: can be have optional ? in front which means its value can be empty string
'-- Lvl2Nm-?-Fst-term-of- expression: It must belong with ?
'-- ?: namepace-? is for 
'-- Output: a hash of all name with namespace
'-- FirstLvl1: first Lvl1 is consider as the output
'Sql
'    Drp  .Drp Tx TxMbr MbrDta Div Sto Crd Cnt Oup MbrWs
'    T    @ Tx UpdTx TxMbr ?MbrDta Div Sto Crd
'    O    @ Cnt Oup ?MbrWs
'Sql.T
'    Tx
'        .Sel@ Crd Amt Qty Cnt ?Mbr ?Div ?Sto ?Dte
'        .Into
'        .Fm SalesHistory
'        .Wh SHSDate between '@P.Fm' and '@P.To'
'        .And@ ?Div ?Sto ?Dte
'        .Gp@ Crd ?Mbr ?Div ?Sto ?Dte
'    UpdTx  .Upd .Set@ TxWD
'    TxMbr
'        .SelDis Mbr
'        .Into
'        .Fm #Tx
'    ?MbrDta ?BrkMbr 
'        .Sel@ Mbr Age Sex Sts Dist Area 
'        .Into
'        .Fm JCMember
'        .Wh JCMCode (Select Mbr From #TxMbr)
'    Div .Sel@ Div Nm Seq Sts .Fm Division
'    Sto .Sel@ Sto Nm CNm .Fm LocTbl
'    Crd .Sel@ Crd Nm .Fm JR_FrqMbrLis_#CrdTy()
'Sql.O
'    Cnt
'        .Sel@ ?MbrCnt RecCnt TxCnt Qty Amt
'        .Into
'        .Fm #Tx
'    Oup
'        .Sel@ Crd ?Mbr ?Sto ?Div ?Dte Amt Qty TxCnt
'        .Into
'        .Fm #Tx x
'        .Jn@ Crd ?Div ?Sto ?Mbr
'    ?MbrWs ?SelMbr 
'        .Sel@ Mbr ?Nm ?Adr ?Mail ?Phone 
'        .Into
'        .Fm JCMember 
'        .Wh JCMCode in (Select Mbr From #TxMbr)
'Sql.T.Tx.Set
'    TxWD ...
'Sql.T.Tx.Sel
'    Crd @ CasewhenThen Else End
'    Amt Sum(SHAmount)
'    Qty Sum(SHQty)
'    Cnt Count(SHInvoice+SHSDate+SHRef)
'    ?Mbr ?BrkMbr JCMMCode
'    ?Div ?BrkDiv @Expr
'    ?Sto ?BrkSto @Expr
'    ?Dte @Expr
'@Sql.T.Tx.Sel.Crd
'    CasewhenThen ...
'    Else ...
'    :NEnd .Repeat :N END~
'    End | :NEnd
'@Sql.T.Tx.And
'    ?Div ?SelDiv And @Fld in (@In)
'    ?Crd ?SelCrd And @Fld in (@In)
'    ?Sto ?SelSto And @Fld in (@In)
'Sql.T.Tx.And.?Div Fld @Expr.Div
'Sql.T.Tx.And.?Crd Fld @Expr.Crd
'Sql.T.Tx.And.?Sto Fld @Expr.Sto
'Sql.T.Tx.And.?Div List @In.Div
'Sql.T.Tx.And.?Sto List @In.Sto
'Sql.T.Tx.And.?Crd List @In.Crd
'Sql.T.Tx.Gp
'    Crd @Expr.Crd
'    ?Mbr ?BrkMbr SHMCode
'    ?Div ?BrkDiv @Expr.Div
'    ?Sto ?BrkSto @Expr.Sto
'    ?Dte ?BrkDte @Expr.Dte
'Expr
'    Div
'    Sto
'    Dte
'        ?SumY @Expr. TxY
'        ?SumM @Expr. TxY TxM
'        ?SumW @Expr. TxY TxM TxW
'        ?SumD @Expr. TxY TxM TxW TxD TxWD TxDte
'    TxY
'    TxM
'    TxW
'    TxD
'    TxWD
'    TxDte
'Sql.T.Print MbrDta.Sel
'    Mbr JCMCode
'    Age DATEDIFF(YEAR,CONVERT(DATETIME ,JCMDOB,112),GETDATE())
'    Sex JCMSex
'    Sts JCMStatus
'    Dist JCMDist
'    Area JCMArea
'Sql.T.Div.Sel
'    Div Dept + Division
'    Nm LongDesc
'    Seq Seq
'    Sts Status
'Sql.T.Sto.Sel
'    Sto '0'+Loc_Code
'    Nm Loc_Name
'    CNm Loc_CName
'Sql.T.Crd.Sel
'    Crd CrdTyId
'    Nm CrdTyNm
'Print
'    SelDiv .Ne @P.DivLis *Blank
'    SelCrd .Ne @P.CrdLis *Blank
'    SelSto .Ne @P.StoLis *Blank
'    BrkDiv .Eq @P.BrkSto 1
'    BrkSto .Eq @P.BrkSto 1
'    BrkSto .Eq @P.BrkSto 1
'    Y .Eq @P.SumLvl Y
'    M .Eq @P.SumLvl M
'    W .Eq @P.SumLvl W
'    D .Eq @P.SumLvl D
'    Dte .Or Y M W D
'    AnyMbrInf @P !Or .InclAdr .InclPhone .InclMail
'    Mbr .And BrkMbr AnyMbrInf
'Sql.O.Oup.Sel
'    ?Mbr ?BrkMbr Mbr
'    ?Sto ?BrkSto Sto
'    ?Div ?BrkDiv Div
'Sql.O.Oup.Sel.Dte
'    Y TxY
'    M TxY TxM
'    W TxY TxM TxW
'    D TxY TxM TxW TxD TxWD TxDte
'Sql.O.Cnt.Sel
'    ?MbrCnt 
'    RecCnt Count(*) RecCnt
'    TxCnt Sum(TxCnt) TxCnt
'    Qty Sum(Qty) Qty
'    Amt Sum(Amt) Amt
'Sql.O.Oup.Jn
'    Crd | Left Join #Crd a #Crd on a.Crd=x.Crd
'    ?Mbr ?BrkMbr | Left Join #MbrDta b on a.Mbr = x.Mbr
'    ?Div ?BrkDiv | Left Join #Div on c.Div = x.Div
'    ?Sto ?BrkSto | Left Join #Sto on d.Sto = x.Sto ---aaa

End Sub

Private Function PrmFm$()

End Function

Private Function PrmTo$()

End Function

Private Sub PushX(OAy, I$)
If I = "" Then Exit Sub
Push OAy, I
End Sub

Private Function Sql() As String()

Dim O$()
PushAy O, Sql_Drp
PushAy O, Sql_T
PushAy O, Sql_O
Sql = O
End Function

Private Function Sql_Drp() As String()
Dim O$()
Push O, Sql_Drp_Tx
Push O, Sql_Drp_TxMbr
Push O, Sql_Drp_MbrDta
Push O, Sql_Drp_Div
Push O, Sql_Drp_Sto
Push O, Sql_Drp_Crd
Push O, Sql_Drp_Cnt
Push O, Sql_Drp_Oup
Push O, Sql_Drp_MbrWs
Sql_Drp = O
End Function

Private Function Sql_Drp_Cnt$()

End Function

Private Function Sql_Drp_Crd$()

End Function

Private Function Sql_Drp_Div$()

End Function

Private Function Sql_Drp_MbrDta$()

End Function

Private Function Sql_Drp_MbrWs$()

End Function

Private Function Sql_Drp_Oup$()

End Function

Private Function Sql_Drp_Sto$()

End Function

Private Function Sql_Drp_Tx$()

End Function

Private Function Sql_Drp_TxMbr$()

End Function

Private Function Sql_O() As String()
Dim O$()
Push O, Sql_O_Cnt
Push O, Sql_O_Oup
PushX O, Sql_O_OptMbrWs
Sql_O = O
End Function

Private Function Sql_O_Cnt$()
'    Cnt .Sel@ ?MbrCnt RecCnt TxCnt Qty Amt .Into .Fm #Tx
Dim O$
O = O & Sql_O_Cnt_Sel
O = O & Sql_O_Cnt_Into
O = O & Sql_O_Cnt_Fm
Sql_O_Cnt = RplVBar(O)
End Function

Private Function Sql_O_Cnt_Fm$()
Sql_O_Cnt_Fm = "|  From #Tx"
End Function

Private Function Sql_O_Cnt_Into$()
Sql_O_Cnt_Into = "|  Into #Cnt"
End Function

Private Function Sql_O_Cnt_Sel$()
Dim O$()
PushX O, Sql_O_Cnt_Sel_OptMbrCnt
Push O, Sql_O_Cnt_Sel_RecCnt
Push O, Sql_O_Cnt_Sel_TxCnt
Push O, Sql_O_Cnt_Sel_Qty
Push O, Sql_O_Cnt_Sel_Amt
Sql_O_Cnt_Sel = "Select " & JnComma(AyAddPfx(O, "|      "))
End Function

Private Function Sql_O_Cnt_Sel_Amt$()
Sql_O_Cnt_Sel_Amt = "Sum(Amt) Amt"
End Function

Private Function Sql_O_Cnt_Sel_OptMbrCnt$()

End Function

Private Function Sql_O_Cnt_Sel_Qty$()
Sql_O_Cnt_Sel_Qty = "Sum(Qty) Qty"
End Function

Private Function Sql_O_Cnt_Sel_RecCnt$()
Sql_O_Cnt_Sel_RecCnt = "Count(*) RecCnt"
End Function

Private Function Sql_O_Cnt_Sel_TxCnt$()
Sql_O_Cnt_Sel_TxCnt = "Sum(TxCnt) TxCnt"
End Function

Private Function Sql_O_MbrWs$()

End Function

Private Function Sql_O_OptMbrWs$()

End Function

Private Function Sql_O_Oup$()

End Function

Private Function Sql_O_Oup_Jn_Crd$()
Sql_O_Oup_Jn_Crd = "| Left Join #Crd a #Crd on a.Crd=x.Crd"
End Function

Private Function Sql_O_Oup_Jn_OptDiv$()
If SwtichBrkDiv Then Sql_O_Oup_Jn_OptDiv = "| Left Join #Div c on c.Div = x.Div"
End Function

Private Function Sql_O_Oup_Jn_OptMbr$()
If SwitchBrkMbr Then Sql_O_Oup_Jn_OptMbr = "| left Join #MbrDta b on a.Mbr = x.Mbr"
End Function

Private Function Sql_O_Oup_Jn_OptSto$()
If SwitchBrkSto Then Sql_O_Oup_Jn_OptSto = "| Left Join #Div c on c.Div = x.Div"
End Function

Private Function Sql_T() As String()
Dim O$()
Push O, Sql_T_Tx
Push O, Sql_T_TxMbr
PushX O, Sql_T_OptMbrDta
Push O, Sql_T_Div
Push O, Sql_T_Sto
Push O, Sql_T_Crd
Sql_T = O
End Function

Private Function Sql_T_Crd$()

End Function

Private Function Sql_T_Div$()

End Function

Private Function Sql_T_OptMbrDta$()

End Function

Private Function Sql_T_Sto$()

End Function

Private Function Sql_T_Tx$()
Dim O$
O = O & Sql_T_Tx_Sel
O = O & Sql_T_Tx_Into
O = O & Sql_T_Tx_Fm
O = O & Sql_T_Tx_Wh
O = O & Sql_T_Tx_And
O = O & Sql_T_Tx_Gp
Sql_T_Tx = O
End Function

Private Function Sql_T_Tx_And$()
Dim O$()
PushX O, Sql_T_Tx_And_OptDiv
PushX O, Sql_T_Tx_And_OptSto
PushX O, Sql_T_Tx_And_OptDte
If AyIsEmpty(O) Then Exit Function
Sql_T_Tx_And = Join(AyAddPfx(O, "|  And "))
End Function

Private Function Sql_T_Tx_And_Crd$()
If SwitchSelCrd Then Sql_T_Tx_And_Crd = FmtQQ(" ? in (?)", Sql_T_Tx_And_Crd_Fld, Sql_T_Tx_And_Crd_In)
End Function

Private Function Sql_T_Tx_And_Crd_Fld$()

End Function

Private Function Sql_T_Tx_And_Crd_In$()

End Function

Private Function Sql_T_Tx_And_Div$()
If SwitchSelDiv Then Sql_T_Tx_And_Div = FmtQQ(" ? in (?)", Sql_T_Tx_And_Div_Fld, Sql_T_Tx_And_Div_In)
End Function

Private Function Sql_T_Tx_And_Div_Fld$()

End Function

Private Function Sql_T_Tx_And_Div_In$()

End Function

Private Function Sql_T_Tx_And_OptDiv$()

End Function

Private Function Sql_T_Tx_And_OptDte$()

End Function

Private Function Sql_T_Tx_And_OptSto$()

End Function

Private Function Sql_T_Tx_And_Sto$()
If SwitchSelSto Then Sql_T_Tx_And_Sto = FmtQQ(" ? in (?)", Sql_T_Tx_And_Sto_Fld, Sql_T_Tx_And_Sto_In)
End Function

Private Function Sql_T_Tx_And_Sto_Fld$()

End Function

Private Function Sql_T_Tx_And_Sto_In$()

End Function

Private Function Sql_T_Tx_Fm$()
Sql_T_Tx_Fm = "|  From SalesHistory"
End Function

Private Function Sql_T_Tx_Gp$()

End Function

Private Function Sql_T_Tx_Into$()
Sql_T_Tx_Into = "|  Into #Tx"
End Function

Private Function Sql_T_Tx_Sel$()
Dim O$()
Push O, Sql_T_Tx_Sel_Crd
Push O, Sql_T_Tx_Sel_Amt
Push O, Sql_T_Tx_Sel_Qty
Push O, Sql_T_Tx_Sel_Cnt
PushX O, Sql_T_Tx_Sel_OptMbr
PushX O, Sql_T_Tx_Sel_OptDiv
PushX O, Sql_T_Tx_Sel_OptSto
PushX O, Sql_T_Tx_Sel_OptDte
Sql_T_Tx_Sel = "Select |  " & Join(O, "|  ")
End Function

Private Function Sql_T_Tx_Sel_Amt$()

End Function

Private Function Sql_T_Tx_Sel_Cnt$()

End Function

Private Function Sql_T_Tx_Sel_Crd$()

End Function

Private Function Sql_T_Tx_Sel_OptDiv$()

End Function

Private Function Sql_T_Tx_Sel_OptDte$()

End Function

Private Function Sql_T_Tx_Sel_OptMbr$()

End Function

Private Function Sql_T_Tx_Sel_OptSto$()

End Function

Private Function Sql_T_Tx_Sel_Qty$()

End Function

Private Function Sql_T_Tx_Wh$()
Sql_T_Tx_Wh = FmtQQ("|  Where SHSDate between '?' and '?'", PrmFm, PrmTo)
End Function

Private Function Sql_T_TxMbr$()

End Function

Private Function SwitchBrkMbr() As Boolean

End Function

Private Function SwitchBrkSto() As Boolean

End Function

Private Function SwitchSelCrd() As Boolean

End Function

Private Function SwitchSelDiv() As Boolean

End Function

Private Function SwitchSelSto() As Boolean

End Function

Private Function SwtichBrkDiv() As Boolean

End Function

Private Sub Sql_O_Cnt__Tst()
Debug.Print Sql_O_Cnt
End Sub
Sub Tst()
Sql_O_Cnt__Tst
End Sub
