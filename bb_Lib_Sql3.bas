Attribute VB_Name = "bb_Lib_Sql3"
Option Compare Database
Option Explicit
Private Enum eOp
    'eStr eNbr eFlag eNbrLis eStrLis eFlag are valid only in Ns:Prm
    eBet     ' [.Bet]
    eEq      ' [.EQ]
    eExpAnd  ' [@And] means <Prm> is term list for "Sql-And"
    eExpComma ' [@Comma]
    eExpDrp
    eExpGp
    eExpJn
    eExpLeftJn
    eExpOr  ' [@Or] means <Prm> is term list for "Sql-Or"
    eExpSel  ' [@Sel] means <Prm> is term list for "Sql-Select"
    eExpSelDis  ' [@SelDis] means <Prm> is term list for "Sql-Select-Distinct"
    eExpTerm  ' [@]   means <Prm> is term list for sql-statment
    eExpSet   ' [@Set] means <Prm> is term list which will be expanded into "Set <Term> = <Exp-term>, .."
    eExpWh   ' [@Wh] means <Prm> is term list for "Sql-Where"
    eFixAnd  ' [.And] means <Prm> is fixed str for "Sql-And"
    eFixComma ' [.Comma]
    eFixDrp
    eFixFm      ' [.Fm]  means <Prm> is fixed str for "Sql-From"
    eFixGp
    eFixInto    ' [.Into] means <Prm> is empty for "Sql-Into" using #<Nm> as the into table name
    eFixOr  ' [.Or] means <Prm> is fixed str for "Sql-Or"
    eFixSel  ' [.Sel] means <Prm> is fixed str for "Sql-Select"
    eFixSet  ' [.Set] means <Prm> is fixed str for "Sql-Set" to be expanded as Set <Prm>
    eFixSelDis ' [.SelDis] means <Prm> is fixed str for "Sql-Select-Distinct"
    eFixStr    ' [.]   means <Prm> is fixed str
    eFixUpd    ' [.Upd]  means <Prm> is empty "Sql-Update" to be expanded Update #<Nm>
    eFixWh     ' [.Wh] means <Prm> is fixed str for "Sql-Where"
    eFlag      ' [.Flag]
    eFixLeftJn ' [.LeftJn] means <Prm> is fixed str for "Sql-Left-Join"
    eFixJn     ' [.Jn] means <Prm> is fixed str for "Sql-inner-Join"
    eMac       ' [$] means <Prm> is a macro string ( a template string with {..} to be expand.  Inside {..} is a <Ns>.<Nm>.
    eMacAnd    ' [$And] means <Prm> is a macro-string
    eMacOr     ' [$Or] means <Prm> is a Macro String to be used in Sql-Or
    eNBet    ' [.NBet]
    eNbr     ' [.Nbr] means <Prm> is a number
    eNbrLis  ' [.NbrLis]
    eNe      ' [.NE]
    eStr     ' [.Str] means <Prm> is a string
    eStrLis  ' [.StrLis]
    eUnknown '

End Enum
Private Type L3
    L3 As String     ' [?<Switch>] <OpTy>[<Op>] [<Prm>]
    Switch As String ' Start with ?, but
    SwitchVal As String ' set by Exp_SwitchVal
    OpStr As String
    Op As eOp
    Prm As String    ' RestTerm of L3
End Type
Private Type WrkDr
    Ns As String
    Nm As String
    L3 As L3
    Str As String   ' Rslt
    LinI As Integer
    Done As Boolean ' Done=(Str="")
End Type

Sub AA_Sql__Flow()
'-- Rmk: -- is remark
'-- 3Lvl: alway 3 level
'-- 4spc: Lvl1 has no space, Lvl2 has exactly 4 space and L3 always have 8 space
'-- NoSpcInNm: Lvl2 (name), cannot have space
'-- Lvl1: is namespace, use do to separate
'-- Lvl2: is name.  That means is always under a namespace
'-- Root Ns: fst non remark line is root ns
'-- L3: is Exp_ression
'-- Lvl2Nm-?: can be have optional ? in front which means its value can be empty string
'-- Lvl2Nm-?-Fst-term-of- Exp_ression: It must belong with ?
'-- ?: namepace-? is for 
'-- Output: a hash of all name with namespace
'-- FirstLvl1: first Lvl1 is consider as the output
'Sql
'    Drp  .Drp Tx TxMbr MbrDta Div Sto Crd Cnt Oup MbrWs
'    T    @ Tx TxMbr ?MbrDta Div Sto Crd
'    O    @ Cnt Oup ?MbrWs
'Sql.T
'    Tx
'        .Sel@ Crd Amt Qty Cnt ?Mbr ?Div ?Sto ?Dte
'        .Into
'        .Fm SalesHistory
'        .Wh SHSDate between '@P.Fm' and '@P.To'
'        .And@ ?Div ?Sto ?Dte
'        .Gp@ Crd ?Mbr ?Div ?Sto ?Dte
'    Tx .Upd .Set@ TxWD
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
'    Cnt .Sel@ ?MbrCnt RecCnt TxCnt Qty Amt .Into .Fm #Tx
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
'    ?Div ?BrkDiv @Exp_r
'    ?Sto ?BrkSto @Exp_r
'    ?Dte @Exp_r
'@Sql.T.Tx.Sel.Crd
'    CasewhenThen ...
'    Else ...
'    :NEnd .Repeat :N END~
'    End | :NEnd
'@Sql.T.Tx.And
'    ?Div ?SelDiv And .Fld@ in (.List@)
'    ?Crd ?SelCrd And .Fld@ in (.List@)
'    ?Sto ?SelSto And .Fld@ in (.List@)
'Sql.T.Tx.And.?Div Fld @Exp_r.Div
'Sql.T.Tx.And.?Crd Fld @Exp_r.Crd
'Sql.T.Tx.And.?Sto Fld @Exp_r.Sto
'Sql.T.Tx.And.?Div List @In.Div
'Sql.T.Tx.And.?Sto List @In.Sto
'Sql.T.Tx.And.?Crd List @In.Crd
'Sql.T.Tx.Gp
'    Crd @Exp_r.Crd
'    ?Mbr ?BrkMbr SHMCode
'    ?Div ?BrkDiv @Exp_r.Div
'    ?Sto ?BrkSto @Exp_r.Sto
'    ?Dte ?BrkDte @Exp_r.Dte
'Exp_r
'    Div
'    Sto
'    Dte
'        ?SumY @Exp_r. TxY
'        ?SumM @Exp_r. TxY TxM
'        ?SumW @Exp_r. TxY TxM TxW
'        ?SumD @Exp_r. TxY TxM TxW TxD TxWD TxDte
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
'    MbrCnt 
'    RecCnt Count(*)
'    TxCnt Sum(TxCnt)
'    Qty Sum(Qty)
'    Amt Sum(Amt)
'Sql.O.Oup.Jn
'    Crd | Left Join #Crd a #Crd on a.Crd=x.Crd
'    ?Mbr ?BrkMbr | Left Join #MbrDta b on a.Mbr = x.Mbr
'    ?Div ?BrkDiv | Left Join #Div on c.Div = x.Div
'    ?Sto ?BrkSto | Left Join #Sto on d.Sto = x.Sto ---aaa

End Sub

Sub AA1()
Exp__Tst
End Sub

Sub Main()
Dim A$(), B$()
bb_Lib_Sql3.Sql3_LyDrs A, B
End Sub

Sub Sql3_Edt()
FtBrw ZZSql3_Ft
End Sub

Private Function Er_NotAlwSwitch(Wy() As WrkDr) As Variant()
Dim O()
    Dim J%, S$
    For J = 0 To UBound(Wy)
        With Wy(J).L3
            If .Switch = "" Then GoTo Nxt
            If Op_IsAlwSwitch(.Op) Then GoTo Nxt
            S = FmtQQ("Switch is not allowed in Op[?].  Only these Op are allowed:?", OpStr(.Op), Op_AlwSwitchOpLis$)
            Push O, Array(Wy(J).LinI, S)
        End With
Nxt:
    Next
Er_NotAlwSwitch = O
End Function
Private Function Op_IsAlwSwitch(A As eOp) As Boolean
Op_IsAlwSwitch = AyHas(Op_AlwSwitchOpAy, A)
End Function
Private Function Op_AlwSwitchOpLis$()
Dim Ay$()
Dim I, Op As eOp
For Each I In Op_AlwSwitchOpAy
    Op = I
    Push Ay, OpStr(Op)
Next
Op_AlwSwitchOpLis = JnSpc(AyQuote(Ay, "[]"))
End Function

Private Function Op_AlwSwitchOpAy() As eOp()
Dim O() As eOp, I
For Each I In Array(eOp.eFixFm, eOp.eFixGp, eOp.eExpGp, eOp.eFixInto, eOp.eFixSelDis, eOp.eExpSelDis, eOp.eMac, eOp.eExpTerm, eOp.eExpComma, _
    eOp.eFixLeftJn, eOp.eFixJn, eOp.eExpJn, eOp.eExpLeftJn, _
    eOp.eExpSel, eOp.eFixSel, eOp.eMacAnd, eOp.eMacOr, eOp.eFixStr)
    Push O, I
Next
Op_AlwSwitchOpAy = O
End Function

Private Function Exp_ThoseWithExp_Str$(ExpOp As eOp, Sy$())
Dim O$
If AyIsEmpty(Sy) Then Exit Function
Select Case ExpOp
Case eOp.eMac: O = Join(Sy, "||")
Case eOp.eExpAnd: O = Quote(Join(AyAddPfx(Sy, "|    and ")), "()")
Case eOp.eExpComma: O = JnComma(AyAddPfx(Sy, "|    "))
Case eOp.eExpDrp: O = Join(Sy, "||")
Case eOp.eExpGp: O = "|  Group by" & Join(AyAddPfx(Sy, "|    "))
Case eOp.eExpJn: O = Join(AyAddPfx(Sy, "|  Inner Join "))
Case eOp.eExpLeftJn: O = Join(AyAddPfx(Sy, "|  Left Join "))
Case eOp.eExpOr: O = Quote(Join(AyAddPfx(Sy, "|    or ")), "()")
Case eOp.eExpSel: O = "Select" & JnComma(AyAddPfx(Sy, "|    "))
Case eOp.eExpSelDis:  O = "Select Distinct" & JnComma(AyAddPfx(Sy, "|    "))
Case eOp.eExpSet: O = "Set" & JnComma(AyAddPfx(Sy, "|    "))
Case Else: Stop
End Select
End Function
Private Function Exp_ThoseWithExp(OWy() As WrkDr) As Boolean
'Return true if all done
Dim J%, M As WrkDr
For J = 0 To UBound(OWy)
    M = OWy(J)
    With M
        If .Done Then GoTo Nxt
        If Not Op_IsExp(.L3.Op) Then GoTo Nxt
        With Macro_ExpTermLis(OWy, .Ns, .Nm, .L3.Prm)
            If Not .Som Then GoTo Nxt
            M.Str = Exp_ThoseWithExp_Str(M.L3.Op, .Sy)
            M.Done = True
            Exp_ThoseWithExp = False
        End With
    End With
Nxt:
Next
End Function
Private Function Op_IsExp(A As eOp) As Boolean
Select Case A
Case eOp.eMac, _
    eOp.eExpAnd, _
    eOp.eExpComma, _
    eOp.eExpDrp, _
    eOp.eExpGp, _
    eOp.eExpJn, _
    eOp.eExpOr, _
    eOp.eExpSel, _
    eOp.eExpSelDis, _
    eOp.eExpSet, _
    eOp.eExpTerm
    Op_IsExp = True
End Select
End Function
Function Sql3_LyDrs(Sql3Ly$(), PrmLy$()) As Drs
Dim Ny() As WrkDr:
Ny = Wrk_Dry(Sql3Ly)
Exp Ny
Dim O As Drs
    O.Fny = SplitSpc("Ns Nm Str")
    O.Dry = Sql3_Dry(Ny)
Sql3_LyDrs = O
End Function
Private Function Er_NoPrm(Wy() As WrkDr) As Variant()
Dim J%
For J = 0 To Wrk_UB(Wy)
    If Wy(J).Ns = "Prm" Then Exit Function
Next
Er_NoPrm = Array(Array(0, "Warning: No Prml namespace"))
End Function
Private Function Er_NoSql(Wy() As WrkDr) As Variant()
Dim J%
For J = 0 To Wrk_UB(Wy)
    If Wy(J).Ns = "Sql" Then Exit Function
Next
Er_NoSql = Array(Array(0, "Warning: No Sql namespace"))
End Function

Private Function Er_Dry(Wy() As WrkDr) As Variant()
Dim O()
PushAy O, Er_InvalidOp(Wy)
PushAy O, Er_NotAlwSwitch(Wy)
PushAy O, Er_SwitchNotExist(Wy)
PushAy O, Er_UpdMstHavNamWithPondSign(Wy)
PushAy O, Er_NoPrm(Wy)
PushAy O, Er_NoSql(Wy)
Er_Dry = O
End Function

Private Function Er_InvalidOp(Wy() As WrkDr) As Variant()
Dim J%
Dim O()
For J = 0 To UBound(Wy)
    With Wy(J).L3
        If .Op = eOp.eUnknown Then Push O, Array(J, FmtQQ("Invalid Op[?]", .OpStr))
    End With
Next
Er_InvalidOp = O
End Function

Private Function Macro_ExpTermLis(Wy() As WrkDr, Ns$, Nm$, TermLis) As SyOpt
'Prm is term-list required to be expanded into a str
'Each term, Ns.Nm.Term, will be used to look up from Wy
'Return None is any term cannot be lookup in Wy
'Return Joined string with ||
Dim Ay$(): Ay = SplitLvs(TermLis)
Dim O$(), T, Pfx$, S$
Pfx = Ns & "." & Nm & "."
For Each T In Ay
    S = Pfx & T
    With Macro_Val(Wy, S)
        If Not .Som Then Exit Function
        Push O, .Str
    End With
Next
Macro_ExpTermLis = SomSy(O)
End Function
Private Sub Exp(Wy() As WrkDr)
Exp_Prm Wy
Exp_FixStr Wy
Exp_Switch Wy
Exp_SwitchVal Wy
Exp_FixFm Wy
Exp_FixInto Wy
Exp_FixUpd Wy
Exp_FixWh Wy
Exp_ThoseWithExp Wy
End Sub
Private Sub Exp__Tst()
Sql3_Rmv3DashInFt ZZSql3_Ft
If Sql3_WrtEr(ZZSql3_Ft) Then FtBrw ZZSql3_Ft: Exit Sub
Dim Wy() As WrkDr: Wy = Wrk_Dry(ZZSql3_Ly)
Exp Wy
Wrk_Brw Wy
End Sub
Private Sub Exp_FixStr(OWy() As WrkDr)
Dim M As WrkDr, J%
For J = 0 To UBound(OWy)
    M = OWy(J)
    With M
        If .L3.Op <> eFixStr Then GoTo Nxt
        If .L3.Switch <> "" Then GoTo Nxt
        If .Done Then Stop
        .Done = True            '<==
        .Str = .L3.Prm           '<==
        OWy(J) = M         '<==
    End With
Nxt:
Next
End Sub
Private Sub Exp_FixStr__Tst()
Sql3_Rmv3DashInFt ZZSql3_Ft
If Sql3_WrtEr(ZZSql3_Ft) Then FtBrw ZZSql3_Ft: Exit Sub
Dim Wy() As WrkDr: Wy = Wrk_Dry(ZZSql3_Ly)
Wrk_Brw Wy
Exp_FixStr Wy
Wrk_Brw Wy
End Sub
Private Sub Exp_FixFm__Tst()
Sql3_Rmv3DashInFt ZZSql3_Ft
If Sql3_WrtEr(ZZSql3_Ft) Then FtBrw ZZSql3_Ft: Exit Sub
Dim Wy() As WrkDr: Wy = Wrk_Dry(ZZSql3_Ly)
Wrk_Brw Wy
Exp_FixFm Wy
Wrk_Brw Wy
End Sub

Private Sub Exp_FixFm(OWy() As WrkDr)
Dim J%, M As WrkDr
For J = 0 To Wrk_Sz(OWy) - 1
    M = OWy(J)
    With M
        If .L3.Op <> eFixFm Then GoTo Nxt
        If .Done Then Stop
        If Trim(.L3.Prm) = "" Then Stop
        .Str = "|  From " & .L3.Prm
        .Done = True
        OWy(J) = M
    End With
Nxt:
Next
End Sub

Private Sub Exp_FixInto(OWy() As WrkDr)
Dim J%, M As WrkDr
For J = 0 To Wrk_Sz(OWy) - 1
    M = OWy(J)
    With M
        If .Done Then GoTo Nxt
        If .L3.Op <> eFixInto Then GoTo Nxt
        .Str = "   Into #" & RmvPfx(.Nm, "?")
        .Done = True
        OWy(J) = M
    End With
Nxt:
Next
End Sub
Private Function Er_UpdMstHavNamWithPondSign(Wy() As WrkDr) As Variant()

End Function
Private Sub Exp_FixUpd(OWy() As WrkDr)
Dim J%, M As WrkDr
For J = 0 To UBound(OWy)
    M = OWy(J)
    With M
        If .L3.Op <> eFixUpd Then GoTo Nxt
        If .Done Then Stop
        .Str = "Update #" & Brk(.Nm, "#").S1
        .Done = True
        OWy(J) = M
    End With
Nxt:
Next
End Sub

Private Sub Exp_FixUpd__Tst()

Sql3_Rmv3DashInFt ZZSql3_Ft
If Sql3_WrtEr(ZZSql3_Ft) Then FtBrw ZZSql3_Ft: Exit Sub
Dim Wy() As WrkDr: Wy = Wrk_Dry(ZZSql3_Ly)
'Wrk_Brw Wy
Exp_FixUpd Wy
Wrk_Brw Wy
End Sub

Private Function L3_Brk(L3$) As L3
Dim L$: L = Trim(L3)
Dim Switch$
    If FstChr(L) = "?" Then
        Switch = RmvFstChr(FstTerm(L))
        L = RmvFstTerm(L)
    End If
Dim OpStr$
    OpStr = FstTerm(L)
   
Dim O As L3
With O
    .L3 = L3
    .Switch = Switch
    .OpStr = OpStr
    .Op = Op(OpStr)
    .Prm = RestTerm(L)
End With
L3_Brk = O
End Function

Private Function L3_OpTy$(L3$)
L3_OpTy = FstChr(FstTerm(L3))
End Function

Private Function L3_Prm$(L3$)
L3_Prm = Brk1(L3, " ").S2
End Function
Function Switch_IsDone(Wy() As WrkDr) As Boolean
Dim J%
For J = 0 To Wrk_UB(Wy)
    If Wy(J).Ns = "?" Then
        If Not Wy(J).Done Then Exit Function
    End If
Next
Switch_IsDone = True
End Function
Sub Exp_Switch__Tst()
Sql3_Rmv3DashInFt ZZSql3_Ft
If Sql3_WrtEr(ZZSql3_Ft) Then FtBrw ZZSql3_Ft: Exit Sub
Dim Wy() As WrkDr: Wy = Wrk_Dry(ZZSql3_Ly)
Wrk_Brw Wy
Exp_Prm Wy
Exp_Switch Wy
Wrk_Brw Wy
End Sub
Function Macro_Ns$(MacroStr$)

End Function
Sub Macro_BrkStr(MacroStr$, ONs$, ONm$)
With BrkRev(MacroStr, ".")
    ONs = RmvFstChr(.S1)
    ONm = RmvLasChr(.S2)
End With
End Sub
Function Macro_Val(Wy() As WrkDr, MacroStr$) As StrOpt
Dim Ns$, Nm$: Macro_BrkStr MacroStr, Ns, Nm
Dim J%
For J = 0 To Wrk_UB(Wy)
    With Wy(J)
        If .Ns <> Ns Then GoTo Nxt
        If .Nm <> Nm Then GoTo Nxt
        If Not .Done Then Exit Function
        Macro_Val = SomStr(.Str)
        Exit Function
    End With
Nxt:
Next
End Function
Sub Exp_Prm(Wy() As WrkDr)
Dim J%, M As WrkDr
For J = 0 To Wrk_UB(Wy)
    M = Wy(J)
    With M
        If .Ns = "Prm" Then
            .Done = True
            .Str = .L3.Prm
            Wy(J) = M
        End If
    End With
Next
End Sub
Private Sub Exp_Prm__Tst()
Sql3_Rmv3DashInFt ZZSql3_Ft
If Sql3_WrtEr(ZZSql3_Ft) Then FtBrw ZZSql3_Ft: Exit Sub
Dim Wy() As WrkDr: Wy = Wrk_Dry(ZZSql3_Ly)
Wrk_Brw Wy
Exp_Prm Wy
Wrk_Brw Wy
End Sub
Private Function Macro_Rpl(Wy() As WrkDr, MacroStr$) As StrOpt
Dim O$
    O = MacroStr
    Dim M$: M = TakBet(O, "{", "}", InclMarker:=True)
    While M <> ""
        With Macro_Val(Wy, M)
            If Not .Som Then Exit Function
            O = Replace(O, M, .Str)
        End With
        M = TakBet(O, "{", "}", InclMarker:=True)
    Wend
Macro_Rpl = SomStr(O)
End Function
Sub Exp_FixWh(Wy() As WrkDr)
Dim J%, M As WrkDr
For J = 0 To UBound(Wy)
    M = Wy(J)
    With M
        If .L3.Op <> eFixWh Then GoTo Nxt
        If .Done Then Stop
        With Macro_Rpl(Wy, .L3.Prm)
            If Not .Som Then GoTo Nxt
            M.Done = True
            M.Str = .Str
            Wy(J) = M
        End With
    End With
Nxt:
Next
End Sub
Private Sub Exp_FixWh__Tst()
Sql3_Rmv3DashInFt ZZSql3_Ft
If Sql3_WrtEr(ZZSql3_Ft) Then FtBrw ZZSql3_Ft: Exit Sub
Dim Wy() As WrkDr: Wy = Wrk_Dry(ZZSql3_Ly)
Exp_Prm Wy
Exp_Switch Wy
'Wrk_Brw Wy
Exp_FixWh Wy
Wrk_Brw Wy
End Sub
Private Sub Exp_ThoseWithExp__Tst()
Sql3_Rmv3DashInFt ZZSql3_Ft
If Sql3_WrtEr(ZZSql3_Ft) Then FtBrw ZZSql3_Ft: Exit Sub
Dim Wy() As WrkDr: Wy = Wrk_Dry(ZZSql3_Ly)
Exp_Prm Wy
Exp_Switch Wy
Exp_SwitchVal Wy
Exp_FixStr Wy
Wrk_Brw Wy
Stop
Exp_ThoseWithExp Wy
Wrk_Brw Wy
End Sub

Private Function Switch_EqNe(Wy() As WrkDr, Prm$, IsEq As Boolean) As BoolOpt
Dim T1$, T2$
    With Brk(Prm, " ")
        T1 = .S1
        T2 = .S2
    End With
Dim V1$
    Dim V1Opt As StrOpt
    V1Opt = Macro_Val(Wy, T1)
    If Not V1Opt.Som Then Exit Function
    V1 = V1Opt.Str
Dim V2$
    If T2 = "*Blank" Then V2 = "" Else V2 = T2
Dim Bool As Boolean
    If IsEq Then
        Bool = V1 = V2
    Else
        Bool = V1 <> V2
    End If
Switch_EqNe = SomBool(Bool)
End Function
Function Switch_Val(Wy() As WrkDr, Switch$) As BoolOpt
Dim J%
For J = 0 To Wrk_UB(Wy)
    With Wy(J)
        If .Ns <> "?" Then GoTo Nxt
        If .Nm <> Switch Then GoTo Nxt
        If Not .Done Then Exit Function
        Switch_Val = SomBool(.Str = "1")
        Exit Function
    End With
Nxt:
Next
End Function
Private Function Switch_Dic(Wy() As WrkDr) As Dictionary
'Return Dic with Switch value is off
Dim J%, O As New Dictionary
For J = 0 To Wrk_UB(Wy)
    With Wy(J)
        If .Ns <> "?" Then GoTo Nxt
        If Not .Done Then Stop
        O.Add .Nm, .Str = "1"
    End With
Nxt:
Next
Set Switch_Dic = O
End Function
Private Sub Exp_SwitchVal__Tst()
Sql3_Rmv3DashInFt ZZSql3_Ft
If Sql3_WrtEr(ZZSql3_Ft) Then FtBrw ZZSql3_Ft: Exit Sub
Dim Wy() As WrkDr: Wy = Wrk_Dry(ZZSql3_Ly)
Exp_Prm Wy
Exp_Switch Wy
Wrk_Brw Wy
Exp_SwitchVal Wy
Wrk_Brw Wy
End Sub
Private Function Switch_Exist(Wy() As WrkDr, Switch$) As Boolean
Dim J%
For J = 0 To UBound(Wy)
    With Wy(J).L3
        If .Switch = "" Then GoTo Nxt
        If .Switch <> Switch Then
            Switch_Exist = True
            Exit Function
        End If
    End With
Nxt:
Next
End Function
Private Function Er_SwitchNotExist(Wy() As WrkDr) As Variant()
Dim J%, O()
For J = 0 To UBound(Wy)
    With Wy(J).L3
        If .Switch = "" Then GoTo Nxt
        If Switch_Exist(Wy, .Switch) Then GoTo Nxt
        Push O, Array(J, "Switch not exist")
    End With
Nxt:
Next
Er_SwitchNotExist = O
End Function
Private Sub Exp_SwitchVal(OWy() As WrkDr)
If Not Switch_IsDone(OWy) Then MsgBox "Exp_SwitchVal is called only after Switch_IsDone": Stop
Dim Dic As Dictionary
Set Dic = Switch_Dic(OWy)

Dim J%, M As WrkDr
For J = 0 To Wrk_UB(OWy)
    M = OWy(J)
    With M.L3
        If .Switch = "" Then GoTo Nxt
        If Not Dic.Exists(.Switch) Then Stop
        M.Done = True
        .SwitchVal = Dic(.Switch)
        M.Str = Dic(.Switch)
        OWy(J) = M
    End With
Nxt:
Next
End Sub
Private Function Switch_TermVal(Wy() As WrkDr, Term$) As BoolOpt
If FstChr(Term) = "?" Then
    Switch_TermVal = Switch_Val(Wy, RmvFstChr(Term))
    Exit Function
End If
If FstChr(Term) = "{" And LasChr(Term) = "}" Then
    With Macro_Val(Wy, Term)
        If .Som Then Switch_TermVal = SomBool(.Str = "1")
    End With
    Exit Function
End If
Stop
End Function

Function Switch_AndOr(Wy() As WrkDr, Prm$, IsAnd As Boolean) As BoolOpt
Dim TermAy$(): TermAy = SplitSpc(Prm)
Dim ValAy() As Boolean
ReDim ValAy(UB(TermAy))
Dim J%
For J = 0 To UB(TermAy)
    With Switch_TermVal(Wy, TermAy(J))
        If Not .Som Then Exit Function
        ValAy(J) = .Bool
    End With
Next
Dim Bool As Boolean
    Dim V
    If IsAnd Then
        Bool = True
        For Each V In ValAy
            If V = False Then Bool = False: Exit For
        Next
    Else
        Bool = False
        For Each V In ValAy
            If V = True Then Exit For
        Next
    End If
Switch_AndOr = SomBool(Bool)
End Function
Private Sub Exp_Switch(OWy() As WrkDr)
'Return True is expanded
If Switch_IsDone(OWy) Then Exit Sub
Dim J%, M As WrkDr, V1$, V2$

For J = 0 To Wrk_UB(OWy)
    M = OWy(J)
    With M
        If .Ns <> "?" Then GoTo Nxt
        If .Done Then GoTo Nxt
        Select Case .L3.Op
        Case eEq, eNe
            With Switch_EqNe(OWy, .L3.Prm, .L3.Op = eEq)
                If .Som Then
                    M.Done = True
                    M.Str = IIf(.Bool, "1", "0")
                    OWy(J) = M
                End If
            End With
        Case eFixAnd, eFixOr
            With Switch_AndOr(OWy, .L3.Prm, .L3.Op = eEq)
                If .Som Then
                    M.Done = True
                    M.Str = IIf(.Bool, "1", "0")
                    OWy(J) = M
                End If
            End With
        Case Else: Stop
        End Select
    End With
Nxt:
Next
End Sub
Private Function Lin_IsL1(L) As Boolean
Dim C$
C = FstChr(L)
Lin_IsL1 = IsLetter(C) Or C = "?"
End Function

Private Function Lin_IsL2(L) As Boolean
If Left(L, 4) = Space(4) Then
    Dim C$: C = Mid(L, 5, 1)
    Lin_IsL2 = IsLetter(C) Or C = "?"
End If
End Function

Private Function Lin_IsL3(L) As Boolean
If Left(L, 8) = Space(8) Then
    Lin_IsL3 = Mid(L, 9, 1) <> " "
End If
End Function

Private Function Lin_Lvl(L) As Byte
If Lin_IsL1(L) Then Lin_Lvl = 1: Exit Function
If Lin_IsL2(L) Then Lin_Lvl = 2: Exit Function
If Lin_IsL3(L) Then Lin_Lvl = 3: Exit Function
Lin_Lvl = 99
End Function

Private Function Op(OpStr$) As eOp
Dim O As eOp
Select Case OpStr
Case "$": O = eMac
Case "$And": O = eMacAnd
Case "$Or": O = eExpOr
Case ".": O = eFixStr
Case ".": O = eFixStr
Case ".And": O = eFixAnd
Case ".Bet": O = eBet
Case ".Comma": O = eFixComma
Case ".Drp": O = eFixDrp
Case ".Eq": O = eEq
Case ".Flag": O = eFlag
Case ".Fm": O = eFixFm
Case ".Gp": O = eFixGp
Case ".Into": O = eFixInto
Case ".Jn": O = eFixJn
Case ".LeftJn": O = eFixLeftJn
Case ".Jn": O = eFixJn
Case ".NBet": O = eNBet
Case ".Nbr": O = eNbr
Case ".NbrLis": O = eNbrLis
Case ".Ne": O = eNe
Case ".Or": O = eFixOr
Case ".Sel": O = eFixSel
Case ".SelDis": O = eFixSelDis
Case ".Set": O = eFixSet
Case ".Str": O = eStr
Case ".StrLis": O = eStrLis
Case ".Upd": O = eFixUpd
Case ".Wh": O = eFixWh
Case "@": O = eExpTerm
Case "@Jn": O = eExpJn
Case "@LeftJn": O = eExpLeftJn
Case "@And": O = eExpAnd
Case "@Comma": O = eExpComma
Case "@Drp": O = eExpDrp
Case "@Gp": O = eExpGp
Case "@Jn": O = eExpJn
Case "@Or": O = eExpOr
Case "@Sel": O = eExpSel
Case "@SelDis": O = eExpSelDis
Case "@Set": O = eExpSet
Case "@Wh": O = eExpWh
Case Else: O = eUnknown
End Select
Op = O
End Function

Private Function Op_Chk$(OpStr$)

End Function

Private Function Op_IsVdt(OpStr$) As Boolean

End Function

Private Function Op_Sy() As String()
Dim O$(), J&
For J = 0 To EnmMbrCnt("eOp", "bb_Lib_Sql3") - 1
    Push O, OpStr(J)
Next
Op_Sy = O
End Function

Private Function OpStr(A As eOp)
Dim O$
Select Case A
Case eBet:   O = ".Bet"
Case eEq:   O = ".Eq"
Case eExpAnd: O = "@And"
Case eExpComma:   O = "@Comma"
Case eExpDrp: O = "@Drp"
Case eExpGp: O = "@Gp"
Case eExpJn: O = "@Jn"
Case eExpLeftJn: O = "@LeftJn"
Case eExpOr: O = "@Or"
Case eExpSel: O = "@Sel"
Case eExpSelDis: O = "@SelDis"
Case eExpSet: O = "@Set"
Case eExpTerm:   O = "@"
Case eExpWh: O = "@Wh"
Case eFixAnd: O = ".And"
Case eFixComma:   O = ".Comma"
Case eFixDrp: O = ".Drp"
Case eFixJn: O = ".Jn"
Case eFixFm: O = ".Fm"
Case eFixGp: O = ".Gp"
Case eFixInto: O = ".Into"
Case eFixJn: O = ".Jn"
Case eFixOr: O = ".Or"
Case eFixSel: O = ".Sel"
Case eFixSelDis: O = ".SelDis"
Case eFixSet: O = ".Set"
Case eFixStr: O = "."
Case eFixUpd: O = ".Upd"
Case eFixWh: O = ".Wh"
Case eFlag: O = ".Flag"
Case eFixLeftJn:   O = ".LeftJn"
Case eMac:   O = "$"
Case eMacAnd: O = "$And"
Case eMacOr:   O = "$Or"
Case eNBet:   O = ".NBet"
Case eNbr: O = ".Nbr"
Case eNbrLis: O = ".NbrLis"
Case eNe: O = ".Ne"
Case eStr: O = ".Str"
Case eStrLis: O = ".StrLis"
Case Else: O = "?Unknown"
End Select
OpStr = O
End Function

Private Function Sql3_Dic(Wy() As WrkDr) As Dictionary
If Wrk_IsEmpty(Wy) Then Exit Function
Dim J%
Dim ONy() As WrkDr
    ONy = Wy
Dim U%
    U = UBound(ONy)
Dim O As New Dictionary
    Dim Exp_uated As Boolean
    Dim N%
    N = 0
    Exp_uated = False
    While Not Exp_uated
'        Exp_uated = Exp_OneCycle(O, ONy)
        N = N + 1
        If N > 1000 Then Stop
    Wend
Set Sql3_Dic = O
End Function

Private Function Sql3_Dry(Wy() As WrkDr) As Variant()
'Dr = Ns Nm Str
Dim Dic As Dictionary
Set Dic = Sql3_Dic(Wy)
Dim ODry(), Dr()
Dim Wrk_ As WrkDr
Dim Nm$, Ns$, Str$, J%
For J = 0 To UBound(Wy)
    Wrk_ = Wy(J)
    With Wrk_
        If Not .Done Then Stop
        Dr = Array(.Ns, .Nm, .Str)
        Push ODry, Dr
    End With
Next
Sql3_Dry = ODry
End Function

Private Sub Sql3_LyBrw(Sql3_Ly$())
Dim Dry(), Dr
Dim L
For Each L In Sql3_Ly
    Push Dry, Array(Lin_Lvl(L), L)
Next
Dim O As Drs
    O.Fny = SplitSpc("Lvl Str")
    O.Dry = Dry
DrsBrw O
End Sub

Private Function Sql3_Rmv2Dash(Ly$()) As String()
Dim O$(), I
For Each I In Ly
    Push O, Brk1(I, "--", NoTrim:=True).S1
Next
Sql3_Rmv2Dash = O
End Function

Private Function Sql3_Rmv3Dash(Ly$()) As String()
Dim O$(), I
For Each I In Ly
    Push O, Brk1(I, "---", NoTrim:=True).S1
Next
Sql3_Rmv3Dash = O
End Function
Private Sub Sql3_Rmv3DashInFt(Ft)
Dim Ly$(): Ly = FtLy(Ft)
Dim Ly1$(): Ly1 = Sql3_Rmv3Dash(Ly)
If AyIsEq(Ly, Ly1) Then Exit Sub
AyWrt Ly1, Ft
End Sub
Private Function Sql3_ValidatedLy(No3Dash_Sql3Ly$()) As String()
Dim O$(): O = No3Dash_Sql3Ly
Dim ErDry(): ErDry = Er_Dry(Wrk_Dry(No3Dash_Sql3Ly))
If AyIsEmpty(ErDry) Then Exit Function
Dim I&, Dr
For Each Dr In ErDry
    I = Dr(0)
    O(I) = O(I) & " --- " & Dr(1)
Next
AyRmvEmptyEleAtEnd O
Push O, FmtQQ("--- [?] error(s)", Sz(ErDry))
Sql3_ValidatedLy = O
End Function

Private Function Wrk_Drs(Wy() As WrkDr) As Drs
Dim ODry(), J%
For J = 0 To UBound(Wy)
    With Wy(J)
        Push ODry, Array(.LinI, .Ns, IIf(.Done, "*", ""), .Nm, .L3.Switch, .L3.SwitchVal, OpStr(.L3.Op), .L3.Prm, .L3.L3, .Str)
    End With
Next
Dim O As Drs
O.Fny = SplitSpc("LinI Ns Done Nm Switch SwitchVal Op Prm L3 Str")
O.Dry = ODry
Wrk_Drs = O
End Function

Private Function Wrk_Dry(Sql3_Ly$()) As WrkDr()
Dim O() As WrkDr
    Dim L, LasNs$, LasNm$, LinI%, L3$
    For Each L In Sql3_Ly
        If Lin_IsL1(L) Then
            LasNs = Trim(L)
        ElseIf Lin_IsL2(L) Then
            With Brk1(Trim(L), " ")
                If .S2 = "" Then
                    LasNm = .S1
                Else
                    L3 = .S2
                    LasNm = .S1
                    GoSub Add
                End If
            End With
        ElseIf Lin_IsL3(L) Then
            L3 = Trim(L)
            GoSub Add
        End If
        LinI = LinI + 1
    Next
Wrk_Dry = O
Exit Function
Add:
    Dim Nm As WrkDr, N%
    With Nm
        .L3 = L3_Brk(L3)
        If .L3.Op = eUnknown Then MsgBox "Er validation is no good"
        .LinI = LinI
        .Nm = LasNm
        .Ns = LasNs
    End With
    ReDim Preserve O(N)
    O(N) = Nm
    N = N + 1
    Return
End Function

Private Function Wrk_DryIsEmpty(A() As WrkDr) As Boolean
On Error Resume Next
Wrk_DryIsEmpty = UBound(A) = -1
Exit Function
Wrk_DryIsEmpty = True
End Function

Private Function Wrk_IsEmpty(A() As WrkDr) As Boolean
Wrk_IsEmpty = Wrk_Sz(A) = 0
End Function

Private Function Wrk_IsDone(A() As WrkDr) As Boolean
Dim J%
For J = 0 To UBound(A)
    If Not A(J).Done Then Exit Function
Next
Wrk_IsDone = True
End Function

Private Sub Wrk_Push(OAy() As WrkDr, M As WrkDr)
Dim N%: N = Wrk_Sz(OAy)
ReDim Preserve OAy(N)
OAy(N) = M
End Sub

Private Sub Wrk_PushAy(OAy() As WrkDr, Ay() As WrkDr)
Dim J%
For J = 0 To Wrk_Sz(Ay) - 1
    Wrk_Push OAy, Ay(J)
Next
End Sub
Private Function Wrk_UB&(Wy() As WrkDr)
Wrk_UB = Wrk_Sz(Wy) - 1
End Function

Private Function Wrk_Sz&(Wy() As WrkDr)
On Error Resume Next
Wrk_Sz = UBound(Wy) + 1
End Function

Private Function ZZSql3_Ft$()
ZZSql3_Ft = TstResPth & "SalRpt.Sql3"
End Function

Private Sub ZZSql3_FtFix()
Dim O$(): O = ZZSql3_Ly
Dim J%
For J = 0 To UB(O)
    O(J) = Replace(O(J), Chr(160), " ")
Next
AyWrt O, ZZSql3_Ft
End Sub

Private Function ZZSql3_Ly() As String()
ZZSql3_Ly = FtLy(ZZSql3_Ft)
End Function

Private Function ZZWrk_Dry() As WrkDr()
ZZWrk_Dry = Wrk_Dry(ZZSql3_Ly)
End Function

Private Sub Op_Sy__Tst()
AyDmp Op_Sy
End Sub

Private Sub Sql3_LyBrw__Tst()
Sql3_LyBrw ZZSql3_Ly
End Sub

Private Sub Sql3_ValidatedLy__Tst()
Dim Ly$(): Ly = Sql3_ValidatedLy(ZZSql3_Ly)
If AyIsEmpty(Ly) Then Exit Sub
AyWrt Ly, ZZSql3_Ft
Sql3_Edt
End Sub
Private Sub Sql3_WrtEr__Tst()
If Sql3_WrtEr(ZZSql3_Ft) Then Sql3_Edt
End Sub

Private Function Sql3_WrtEr(Ft) As Boolean
Dim Ly$(): Ly = FtLy(Ft)
Dim Ly1$(): Ly1 = Sql3_ValidatedLy(Ly): If AyIsEmpty(Ly1) Then Exit Function
If AyIsEq(Ly, Ly1) Then Exit Function
AyWrt Ly1, Ft
Sql3_WrtEr = True
End Function
Private Sub Wrk_Brw(Wy() As WrkDr)
DrsWs Wrk_Drs(Wy)
End Sub
Private Sub Wrk_Dry__Tst()
Sql3_Rmv3DashInFt ZZSql3_Ft
If Sql3_WrtEr(ZZSql3_Ft) Then FtBrw ZZSql3_Ft: Exit Sub
Dim Act() As WrkDr: Act = Wrk_Dry(ZZSql3_Ly)
Wrk_Brw Act
Sql3_LyBrw__Tst
End Sub
