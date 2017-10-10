Attribute VB_Name = "SalRpt"
Option Compare Database
Option Explicit

Sub AA1()
Debug.Print SR_Sql
End Sub

Sub SalRpt_Stop()
Stop
End Sub

Function SR_Sql$(Optional SRPNm$)
Dim P As SRP
    P = SRP(SRPNm)
Dim BrkMbr As Boolean
Dim InclNm As Boolean
Dim InclAdr As Boolean
Dim InclEmail As Boolean
Dim InclPhone As Boolean
    InclNm = P.InclNm
    InclAdr = P.InclAdr
    InclEmail = P.InclEmail
    InclPhone = P.InclPhone
Dim O$(), ECrd$
    ECrd = SRECrd(P.CrdLis, SRCrdPfxTyDry)
Push O, ZMulSql_Drp
Push O, SR_SqlT(P, ECrd)
Push O, SR_SqlO(BrkMbr, InclNm, InclAdr, InclEmail, InclPhone)
SR_Sql = RplVbar(JnCrLf(O))
End Function

Function SR_SqlOCnt$()

End Function

Function SR_SqlOMbrWsOpt$( _
    BrkMbr As Boolean, _
    InclNm As Boolean, _
    InclAdr As Boolean, _
    InclEmail As Boolean, _
    InclPhone As Boolean)
Const SR_ETMbrWsOptEAge$ = "DateDiff(Year, Convert(DateTime, JCMDOB, 112), GETDATE())"
Const SR_ETMbrWsOptEMbr$ = "JCMCode"
Const ESex$ = "JCMSex"
Const ESts$ = "JCMStatus"
Const EDist$ = "JCMDist"
Const EArea$ = "JCMArea"
Const EAdr$ = "Adr-Express-L1|Adr-Expression-L2"
Const ENm$ = "JCMName"
Const ECNm$ = "JCMCName"
Const EPhone$ = "JCMPhone"
Const EEmail$ = "JCMEmail"

'Sql.X.T.Print MbrDta
'    Sel # Mbr Age Sex Sts Dist Area ?Nm ?Email ?Phone ?Adr
'    Fm  # JCMember
'    Wh  # JCMCode (Select Mbr From #TxMbr)
'Sql.X.T.Print MbrDta.Sel
'    Mbr .JCMCode
'    Age .DateDiff(Year, Convert(DateTime, JCMDOB, 112), GETDATE())
'    Sex .JCMSex
'    Sts .JCMStatus
'    Dist .JCMDist
'    Area .JCMArea
If Not BrkMbr Then Exit Function
Dim Fny$()
    Dim Ay$()
    Ay = SplitLvs("Mbr Age Sex Sts Dist Area")
    If InclNm Then Push Ay, "Nm"
    If InclAdr Then Push Ay, "Adr"
    If InclEmail Then Push Ay, "Email"
    If InclPhone Then Push Ay, "Phone"
    Fny = Ay
Dim ExprAy()
    Dim Dic As New Dictionary
    With Dic
        .Add "Mbr", SR_ETMbrWsOptEMbr
        .Add "Age", SR_ETMbrWsOptEAge
        .Add "Sex", ESex
        .Add "Sts", ESts
        .Add "Dist", EDist
        .Add "Area", EArea
        If InclAdr Then .Add "Adr", EAdr
        If InclNm Then .Add "Nm", ENm
        If InclNm Then .Add "CNm", ECNm
        If InclPhone Then .Add "Phone", EPhone
        If InclEmail Then .Add "Email", EEmail
    End With
    ExprAy = DicSelIntoAy(Dic, Fny)
Dim Sel$, Into$, Fm$, Wh$
    Sel = SqpSel(Fny, ExprAy)
    Into = SqpInto("#MbrDta")
    Fm = SqpFm("JCMember")
    Wh = SqpWh("JCMDCode in (Select Mbr From #TxMbr)")
SR_SqlOMbrWsOpt = Sel & Into & Fm & Wh
End Function

Function SR_SqlOOup$()
Const L$ = _
"Sel " & _
"|Into #Oup" & _
"|Fm #Tx x" & _
"|Left #TxMbr a on x.Mbr = a.JCMMCode"
SR_SqlOOup = RplVbar(L)
End Function

Function SR_SqlTCrd$(BrkCrd As Boolean, InCrd$)
Const FldLvs$ = "Crd CrdNm"
Const ECrd$ = "CrdTyId"
Const ECrdNm$ = "CrdTyNm"
'Sql.X.T.Crd
'    Sel  # Crd Nm
'    Fm   # JR_FrqMbrLis_#CrdTy()
'Sql.X.T.Crd.Sel
'    Crd .CrdTyId
'    Nm .CrdTyNm
If Not BrkCrd Then Exit Function
Dim ExprAy$()
    ExprAy = Sy(ECrd, ECrdNm)
Dim Sel$, Into$, Fm$, Wh$
    Sel = SqpSelFldLvs(FldLvs, ExprAy)
    Into = SqpInto("#Crd")
    Fm = SqpFm("JR_FrqMbrLis_#CrdTy()")
    Wh = IIf(InCrd = "", "", SqpWh(FmtQQ("? in (?)", ECrd, InCrd)))
SR_SqlTCrd = Sel & Into & Fm & Wh
End Function

Function SR_SqlTDiv$(BrkDiv As Boolean, InDiv$)
If Not BrkDiv Then Exit Function
Const FldLvs$ = "Div DivNm DivSeq DivSts"
Const EDiv$ = "Dept + Division"
Const EDivNm$ = "DivNm"
Const EDivSeq$ = "Seq"
Const EDivSts$ = "Status"
Dim ExprAy$()
    ExprAy = Sy(EDiv, EDivNm, EDivSeq, EDivSts)

Dim Sel$, Into$, Fm$, Wh$
    Sel = SqpSelFldLvs(FldLvs, ExprAy)
    Into = SqpInto("#Div")
    Fm = SqpFm("Division")
    Wh = "": If InDiv <> "" Then Wh = SqpWh(FmtQQ("? in (?)", EDiv, InDiv))
SR_SqlTDiv = Sel & Into & Fm & Wh
End Function

Function SR_SqlTSto$(BrkSto As Boolean, InSto$)
'Sql.X.T.Sto
'    Sel  # Sto Nm CNm
'    Fm   # LocTbl
'Sql.X.T.Sto.Sel
'    Sto . '0'+Loc_Code
'    Nm .Loc_Name
'    CNm .Loc_CName
If Not BrkSto Then Exit Function
Const ESto$ = "'0'+Loc_Code"
Const EStoNm$ = "Loc_Name"
Const EStoCNm$ = "Loc_CName"
Dim ExprAy$()
    ExprAy = Sy(ESto, EStoNm, EStoCNm)
Dim Sel$, Into$, Fm$, Wh$
    Sel = SqpSelFldLvs("Sto StoNm StoCNm", ExprAy)
    Into = SqpInto("#Sto")
    Fm = SqpFm("Location")
    Wh = IIf(InSto = "", "", SqpWh(FmtQQ("? in (?)", ESto, InSto)))
SR_SqlTSto = Sel & Into & Fm & Wh
End Function

Function SR_SqlTTx$(P As SRP, ECrd$)
Dim O$()
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
Dim PX As SRPX
    PX = SRPX(P, ECrd)
Dim Fny$()
    Erase O
    Push O, "Crd"
    Push O, "Amt"
    Push O, "Qty"
    Push O, "Cnt"
    If P.BrkMbr Then Push O, "Mbr"
    If P.BrkDiv Then Push O, "Div"
    If P.BrkSto Then Push O, "Sto"
    If PX.InclFldTxY Then Push O, "TxY"
    If PX.InclFldTxM Then Push O, "TxM"
    If PX.InclFldTxW Then Push O, "TxW"
    If PX.InclFldTxD Then Push O, "TxD"
    If PX.InclFldTxD Then Push O, "TxDte"
    Fny = O
    
Dim ExprAy$()
    Erase O
    Push O, ECrd
    Push O, EAmt
    Push O, EQty
    Push O, ECnt
    If P.BrkMbr Then Push O, EMbr
    If P.BrkDiv Then Push O, EDiv
    If P.BrkSto Then Push O, ESto
    If PX.InclFldTxY Then Push O, ETxY
    If PX.InclFldTxM Then Push O, ETxM
    If PX.InclFldTxW Then Push O, ETxW
    If PX.InclFldTxD Then Push O, ETxD
    If PX.InclFldTxD Then Push O, ETxDte
    ExprAy = O
    
Dim CrdIn$, DivIn$, StoIn$
    CrdIn = SqpExprIn(ECrd, PX.InCrd)
    DivIn = SqpExprIn(EDiv, PX.InDiv)
    StoIn = SqpExprIn(ESto, PX.InSto)
    
Dim GpExprAy$()
    Erase O
    Push O, ECrd
    If P.BrkMbr Then Push O, EMbr
    If P.BrkDiv Then Push O, EDiv
    If P.BrkSto Then Push O, ESto
    If PX.InclFldTxY Then Push O, ETxY
    If PX.InclFldTxM Then Push O, ETxM
    If PX.InclFldTxD Then Push O, ETxD
    If PX.InclFldTxD Then Push O, ETxDte
    GpExprAy = O

Dim Sel$, Into$, Fm$, Wh$, AndCrd$, AndSto$, AndDiv$, Gp$
    Sel = SqpSel(Fny, ExprAy)
    Into = SqpInto("#Tx")
    Fm = SqpFm("SaleHistory")
    Wh = SqpWhBetStr("SHDate", P.FmDte, P.ToDte)
    AndCrd = SqpAnd(CrdIn)
    AndDiv = SqpAnd(DivIn)
    AndSto = SqpAnd(StoIn)
    Gp = SqpGp(GpExprAy)
SR_SqlTTx = Sel & Into & Fm & Wh & AndCrd & AndDiv & AndSto & Gp
End Function

Function SR_SqlTTxMbr$(BrkMbr As Boolean)
If Not BrkMbr Then Exit Function
SR_SqlTTxMbr = "Select Distinct Mbr From #Tx Into #TxMbr"
End Function

Function SR_SqlTTxSel$(Fny$(), ExprAy$())
SR_SqlTTxSel = SqpSel(Fny, ExprAy)
End Function

Function SR_SqlTUpdTx$(InclFldTxDte As Boolean)
Const SR_TUpdTxETxWD$ = _
"CASE WHEN TxWD1 = 1 then 'Sun'" & _
"|ELSE WHEN TxWD1 = 2 THEN 'Mon'" & _
"|ELSE WHEN TxWD1 = 3 THEN 'Tue'" & _
"|ELSE WHEN TxWD1 = 4 THEN 'Mon'" & _
"|ELSE WHEN TxWD1 = 5 THEN 'Thu'" & _
"|ELSE WHEN TxWD1 = 6 THEN 'Fri'" & _
"|ELSE WHEN TxWD1 = 7 THEN 'Sat'" & _
"|ELSE Null" & _
"|END END END END END END END"
If Not InclFldTxDte Then Exit Function
SR_SqlTUpdTx = SqpUpd("#Tx") & SqpSet("TxWD", Array(SR_TUpdTxETxWD))
End Function

Function SRCrdPfxTyDry() As Variant()
Static X As Boolean, Y
If Not X Then
    X = True
    Dim O()
    Push O, Array("134234", 1)
    Push O, Array("12323", 1)
    Push O, Array("2444", 2)
    Push O, Array("2443434", 2)
    Push O, Array("24424", 2)
    Push O, Array("3", 3)
    Push O, Array("5446561", 4)
    Push O, Array("6234341", 5)
    Push O, Array("6234342", 5)
    Y = O
End If
SRCrdPfxTyDry = Y
End Function

Private Function SR_SqlO$( _
    BrkMbr As Boolean, _
    InclNm As Boolean, _
    InclAdr As Boolean, _
    InclEmail As Boolean, _
    InclPhone As Boolean)
Dim O$()
Push O, SR_SqlOCnt
Push O, SR_SqlOOup
Push O, SR_SqlOMbrWsOpt(BrkMbr, InclNm, InclAdr, InclEmail, InclEmail)
O = AyRmvEmpty(O)
SR_SqlO = JnDblCrLf(O)
End Function

Private Function SR_SqlT$(P As SRP, ECrd$)
Dim O$()
Dim PX As SRPX
    PX = SRPX(P, ECrd)
With P
Push O, SR_SqlTTx(P, ECrd)
Push O, SR_SqlTUpdTx(PX.InclFldTxD)
Push O, SR_SqlTTxMbr(.BrkMbr)
Push O, SR_SqlTDiv(.BrkDiv, .DivLis)
Push O, SR_SqlTSto(.BrkSto, .StoLis)
Push O, SR_SqlTCrd(.BrkCrd, .CrdLis)
End With
SR_SqlT = JnCrLf(O)
End Function

Private Function ZMulSql_Drp$()
ZMulSql_Drp = MulSqlDrp("#Tx #TxMbr #MbrDta #Div #Sto #Crd #Cnt #Oup #MbrWs")
End Function

