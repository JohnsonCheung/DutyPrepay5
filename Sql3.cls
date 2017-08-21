VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sql3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Private Type NsNm
    Ns As String
    Nm As String
End Type
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
    eFixFm     ' [.Fm]  means <Prm> is fixed str for "Sql-From"
    eFixGp     ' [.Gp]  means <Prm> is fixed str for "Sql-Group"
    eFixInto   ' [.Into] means <Prm> is empty for "Sql-Into" using #<Nm> as the into table name
    eFixOr     ' [.Or] means <Prm> is fixed str for "Sql-Or"
    eFixSel    ' [.Sel] means <Prm> is fixed str for "Sql-Select"
    eFixSet    ' [.Set] means <Prm> is fixed str for "Sql-Set" to be expanded as Set <Prm>
    eFixSelDis ' [.SelDis] means <Prm> is fixed str for "Sql-Select-Distinct"
    eFixStr    ' [.]   means <Prm> is fixed str
    eFixUpd    ' [.Upd]  means <Prm> is empty "Sql-Update" to be expanded Update #<Nm>
    eFixWh     ' [.Wh] means <Prm> is fixed str for "Sql-Where"
    eFlag      ' [.Flag]
    eFixLeftJn ' [.LeftJn] means <Prm> is fixed str for "Sql-Left-Join"
    eFixJn     ' [.Jn] means <Prm> is fixed str for "Sql-inner-Join"
    eMac       ' [$] means <Prm> is a macro string ( a template string with {..} to be expand.  Inside {..} is a <Ns>.<Nm>.
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
    SwitchVal As String ' set by .Exp_OffSwitchVal
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
'Export: Init Sql
Private A_() As String ' Sql3Ly$()
Private Wy() As WrkDr
Private ErDry() As Variant ' Dr = Lno Msg
Private IsOK As Boolean
Private IsEr As Boolean
Property Get Ly() As String()
Ly = A_
End Property
Function Init(Ly() As String) As Sql3
A_ = Ly
Wy = Init_Wy
ErDry = Init_Er_Dry
IsOK = AyIsEmpty(ErDry)
IsEr = Not IsOK
End Function
Friend Property Get Sql$()
Sql = Dic()("Sql")
End Property
Friend Property Get Ky() As String()

End Property

Friend Property Get Dic() As Dictionary

End Property



Sub AA_Sql__Flow()
'-- Rmk: -- is remark
'-- 3Lvl: alway 3 level
'-- 4spc: Lvl1 has no space, Lvl2 has exactly 4 space and L3 always have 8 space
'-- NoSpcInNm: Lvl2 (name), cannot have space
'-- Lvl1: is namespace, use do to separate
'-- Lvl2: is name.  That means is always under a namespace
'-- Root Ns: fst non remark line is root ns
'-- L3: is .Exp_ression
'-- Lvl2Nm-?: can be have optional ? in front which means its value can be empty string
'-- Lvl2Nm-?-Fst-term-of- .Exp_ression: It must belong with ?
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
'    ?Div ?BrkDiv @.Exp_r
'    ?Sto ?BrkSto @.Exp_r
'    ?Dte @.Exp_r
'@Sql.T.Tx.Sel.Crd
'    CasewhenThen ...
'    Else ...
'    :NEnd .Repeat :N END~
'    End | :NEnd
'@Sql.T.Tx.And
'    ?Div ?SelDiv And .Fld@ in (.List@)
'    ?Crd ?SelCrd And .Fld@ in (.List@)
'    ?Sto ?SelSto And .Fld@ in (.List@)
'Sql.T.Tx.And.?Div Fld @.Exp_r.Div
'Sql.T.Tx.And.?Crd Fld @.Exp_r.Crd
'Sql.T.Tx.And.?Sto Fld @.Exp_r.Sto
'Sql.T.Tx.And.?Div List @In.Div
'Sql.T.Tx.And.?Sto List @In.Sto
'Sql.T.Tx.And.?Crd List @In.Crd
'Sql.T.Tx.Gp
'    Crd @.Exp_r.Crd
'    ?Mbr ?BrkMbr SHMCode
'    ?Div ?BrkDiv @.Exp_r.Div
'    ?Sto ?BrkSto @.Exp_r.Sto
'    ?Dte ?BrkDte @.Exp_r.Dte
'.Exp_r
'    Div
'    Sto
'    Dte
'        ?SumY @.Exp_r. TxY
'        ?SumM @.Exp_r. TxY TxM
'        ?SumW @.Exp_r. TxY TxM TxW
'        ?SumD @.Exp_r. TxY TxM TxW TxD TxWD TxDte
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


Sub Main()
Dim A$(), B$()
Sql3_LyDrs A, B
End Sub
Private Function ZZSql3_Ft$()
ZZSql3_Ft = TstResPth & "SalRpt.Sql3"
End Function

Sub ZZEdt()
End Sub

Friend Function Er_NotAlwSwitch() As Variant()
Dim O()
    Dim J%, S$
    For J = 0 To WyUB
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
Friend Function Op_IsAlwSwitch(A As eOp) As Boolean
'Op_IsAlwSwitch = AyHas(Op_AlwSwitchOpAy, A)
End Function
Friend Function Op_AlwSwitchOpLis$()
Dim Ay$()
Dim I, Op As eOp
For Each I In Op_AlwSwitchOpAy
    Op = I
    Push Ay, OpStr(Op)
Next
'Op_AlwSwitchOpLis = JnSpc(AyQuote(Ay, "[]"))
End Function

Friend Function Op_AlwSwitchOpAy() As eOp()
Dim O() As eOp, I
For Each I In Array(eOp.eFixFm, eOp.eFixGp, eOp.eExpGp, eOp.eFixInto, eOp.eFixSelDis, eOp.eExpSelDis, eOp.eMac, eOp.eExpTerm, eOp.eExpComma, _
    eOp.eFixLeftJn, eOp.eFixJn, eOp.eExpJn, eOp.eExpLeftJn, _
    eOp.eExpSel, eOp.eFixSel, eOp.eFixStr)
    Push O, I
Next
Op_AlwSwitchOpAy = O
End Function

Friend Function NmVal__Mix(ExpOp As eOp, Sy$()) As StrOpt
Dim O$
If AyIsEmpty(Sy) Then Exit Function
Select Case ExpOp
Case eOp.eMac: O = Join(Sy, "||")
'Case eOp.eExpAnd: O = Quote(Join(AyAddPfx(Sy, "|    and ")), "()")
'Case eOp.eExpComma: O = JnComma(AyAddPfx(Sy, "|    "))
Case eOp.eExpDrp: O = Join(Sy, "||")
'Case eOp.eExpGp: O = "|  Group by" & Join(AyAddPfx(Sy, "|    "))
'Case eOp.eExpJn: O = Join(AyAddPfx(Sy, "|  Inner Join "))
'Case eOp.eExpLeftJn: O = Join(AyAddPfx(Sy, "|  Left Join "))
'Case eOp.eExpOr: O = Quote(Join(AyAddPfx(Sy, "|    or ")), "()")
'Case eOp.eExpSel: O = "Select" & JnComma(AyAddPfx(Sy, "|    "))
'Case eOp.eExpSelDis:  O = "Select Distinct" & JnComma(AyAddPfx(Sy, "|    "))
'Case eOp.eExpSet: O = "Set" & JnComma(AyAddPfx(Sy, "|    "))
Case Else: Stop
End Select
NmVal__Mix = StrOpt(O)
End Function
Friend Function Exp_ThoseWithExp() As Boolean
'Return true if all done
Dim J%, M As WrkDr
For J = 0 To WyUB
    M = Wy(J)
    With M
        If .Done Then GoTo Nxt
        If Not Op_IsExp(.L3.Op) Then GoTo Nxt
        With NmVal(.L3.Op, .Ns, .Nm, .L3.Prm)
            If Not .Som Then GoTo Nxt
            M.Str = .Str
            M.Done = True
            Exp_ThoseWithExp = False
        End With
    End With
Nxt:
Next
End Function
Friend Function NmVal(Op As eOp, Ns$, Nm$, TermLis$) As StrOpt
Dim NsNmAy() As NsNm: NsNmAy = NmVal__NsNmAy(Ns, Nm, TermLis)
With NmVal__SyOpt(NsNmAy)
    If Not .Som Then Exit Function
    NmVal = NmVal__Mix(Op, .Sy)
End With
End Function
Friend Function Op_IsExp(A As eOp) As Boolean
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
Exp
Dim Fny$(), Dry As Dry
    Fny = SplitSpc("Ns Nm Str")
    Dry = Sql3_Dry
Sql3_LyDrs = Nw.Drs(Fny, Dry)
End Function
Friend Function Er_NoPrm() As Variant()
Dim J%
For J = 0 To WyUB
    If Wy(J).Ns = "Prm" Then Exit Function
Next
Er_NoPrm = Array(Array(0, "Warning: No Prml namespace"))
End Function
Friend Function Er_NoSql() As Variant()
Dim J%
For J = 0 To WyUB
    If Wy(J).Ns = "Sql" Then Exit Function
Next
Er_NoSql = Array(Array(0, "Warning: No Sql namespace"))
End Function

Private Function Init_Er_Dry() As Variant()
Dim O()
PushAy O, Er_InvalidOp
PushAy O, Er_NotAlwSwitch
PushAy O, Er_SwitchNotExist
PushAy O, Er_UpdMstHavNamWithPondSign
PushAy O, Er_NoPrm
PushAy O, Er_NoSql
Init_Er_Dry = O
End Function
Friend Sub Er_Dry__Tst()
'DryBrw ZZ.Er_Dry
End Sub
Private Function Init_Wy() As WrkDr()
Dim O() As WrkDr
    Dim L, LasNs$, LasNm$, LinI%, L3$
    For Each L In A_
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
Init_Wy = O
Exit Function
Add:
    Dim Nm As WrkDr, N%
    With Nm
        .L3 = L3_Brk(L3)
        If .L3.Op = eUnknown Then MsgBox "Er validation is no good giving an invalid Op": Stop
        .LinI = LinI
        .Nm = LasNm
        .Ns = LasNs
    End With
    ReDim Preserve O(N)
    O(N) = Nm
    N = N + 1
    Return

End Function
Friend Function Er_InvalidOp() As Variant()
Dim J%
Dim O()
For J = 0 To WyUB
    With Wy(J).L3
        If .Op = eOp.eUnknown Then Push O, Array(J, FmtQQ("Invalid Op[?]", .OpStr))
    End With
Next
Er_InvalidOp = O
End Function

Friend Function NmVal__SyOpt(NsNmAy() As NsNm) As SyOpt
Dim J%, O$()
For J = 0 To UBound(NsNmAy)
    With Wrk_LookupVal(NsNmAy(J))
        If Not .Som Then Exit Function
        Push O, .Str
    End With
Next
NmVal__SyOpt = SyOpt(O)
End Function
Friend Sub Exp()
If Not IsOK Then Exit Sub
Exp_Prm
Exp_Switch
Exp_OffSwitchVal
Exp_Fix
Dim Aft%, Bef%
Dim J%
For J = 1 To 50
    Bef = Wrk_DoneCnt
    Exp_ThoseWithExp
    Aft = Wrk_DoneCnt
    If Bef = Aft Then Exit For
Next
Stop
Stop
End Sub
Friend Function Wrk_DoneCnt%()
Dim J%, O%
For J = 0 To WyUB
    If Wy(J).Done Then O = O + 1
Next
Wrk_DoneCnt = O
End Function
Friend Sub Exp__Tst()
With ZZ
    .Exp
    .Wrk_Brw
End With
End Sub
Private Function ZZ() As Sql3
Dim M As New Sql3
M.Init ZZSql3_Ly
Set ZZ = M
End Function
Friend Sub Exp_FixStr()
Dim M As WrkDr, J%
For J = 0 To WyUB
    M = Wy(J)
    With M
        If .L3.Op <> eFixStr Then GoTo Nxt
        If .L3.Switch <> "" Then GoTo Nxt
        If .Done Then Stop
        .Done = True            '<==
        .Str = .L3.Prm           '<==
        Wy(J) = M         '<==
    End With
Nxt:
Next
End Sub
Friend Sub Exp_FixStr__Tst()
With ZZ
    .Wrk_Brw
    .Exp_FixStr
    .Wrk_Brw
End With
End Sub
Friend Sub Exp_FixFm__Tst()
With ZZ
    .Wrk_Brw
    .Exp_Fix
    .Wrk_Brw
End With
End Sub
Friend Sub Exp_Fix()
Exp_Fix
Exp_FixInto
Exp_FixUpd
Exp_FixWh
Exp_FixStr
'Exp_FixGp
'Exp_FixAnd
'Exp_FixFm
'Exp_FixJn
'Exp_FixLeftJn
'Exp_FixOr
'Exp_FixSel
'Exp_FixSelDist
End Sub
Friend Sub Exp_Fix_Itm(Pfx$)
Dim J%, M As WrkDr
For J = 0 To WySz - 1
    M = Wy(J)
    With M
        If .L3.Op <> eFixFm Then GoTo Nxt
        If .Done Then Stop
        If Trim(.L3.Prm) = "" Then Stop
        .Str = Pfx & .L3.Prm
        .Done = True
        Wy(J) = M
    End With
Nxt:
Next
End Sub

Friend Sub Exp_FixFm_Original()
Dim J%, M As WrkDr
For J = 0 To WySz - 1
    M = Wy(J)
    With M
        If .L3.Op <> eFixFm Then GoTo Nxt
        If .Done Then Stop
        If Trim(.L3.Prm) = "" Then Stop
        .Str = "|  From " & .L3.Prm
        .Done = True
        Wy(J) = M
    End With
Nxt:
Next
End Sub

Friend Sub Exp_FixInto()
Dim J%, M As WrkDr
For J = 0 To WySz - 1
    M = Wy(J)
    With M
        If .Done Then GoTo Nxt
        If .L3.Op <> eFixInto Then GoTo Nxt
        .Str = "|   Into #" & RmvPfx(.Nm, "?")
        .Done = True
        Wy(J) = M
    End With
Nxt:
Next
End Sub
Friend Function Er_UpdMstHavNamWithPondSign() As Variant()

End Function
Friend Sub Exp_FixUpd()
Dim J%, M As WrkDr
For J = 0 To WyUB
    M = Wy(J)
    With M
        If .L3.Op <> eFixUpd Then GoTo Nxt
        If .Done Then Stop
        .Str = "Update #" & Brk(.Nm, "#").S1
        .Done = True
        Wy(J) = M
    End With
Nxt:
Next
End Sub

Friend Sub Exp_FixUpd__Tst()
With ZZ
    '.Wrk_Brw
    .Exp_FixUpd
    .Wrk_Brw
End With
End Sub

Friend Function L3_Brk(L3$) As L3
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

Friend Function L3_OpTy$(L3$)
L3_OpTy = FstChr(FstTerm(L3))
End Function

Friend Function L3_Prm$(L3$)
L3_Prm = Brk1(L3, " ").S2
End Function
Friend Property Get Switch_IsDone() As Boolean
'This function is to test if all the Switch (Ns=?) should be Done.
'
Dim J%
For J = 0 To WyUB
    If Wy(J).Ns = "?" Then
        If Not Wy(J).Done Then Exit Property
    End If
Next
Switch_IsDone = True
End Property
Sub Exp_Switch__Tst()
With ZZ
    .Wrk_Brw
    .Exp_Prm
    .Exp_Switch
    .Wrk_Brw
End With
End Sub
Function Macro_Ns$(MacroStr$)

End Function

Friend Function NsNm_ByMacroTerm(MacroTerm$) As NsNm
Dim O As NsNm
With BrkRev(MacroTerm, ".")
    O.Ns = RmvFstChr(.S1)
    O.Nm = RmvLasChr(.S2)
End With
NsNm_ByMacroTerm = O
End Function
Friend Function NmVal__NsNmAy(Ns$, Nm$, TermLis$) As NsNm()
Dim Ay$(): Ay = SplitSpc(TermLis)
Dim O() As NsNm, J%
ReDim O(UB(Ay))
For J = 0 To UB(Ay)
    O(J) = NmVal__NsNm(Ns, Nm, Ay(J))
Next
NmVal__NsNmAy = O
End Function
Friend Function NsNm(Ns$, Nm$) As NsNm
NsNm.Nm = Nm
NsNm.Ns = Ns
End Function
Friend Function NmVal_IsMacro(Term$) As Boolean
If FstChr(Term) = "{" Then
    If LasChr(Term) = "}" Then NmVal_IsMacro = True
    Stop
End If
End Function
Friend Function NmVal__NsNm(Ns$, Nm$, Term$) As NsNm
If NmVal_IsMacro(Term) Then
    NmVal__NsNm = NsNm_ByMacroTerm(Term)
Else
    NmVal__NsNm = NsNm(Ns & "." & Nm, Term)
End If
End Function

Friend Function Wrk_LookupVal(Itm As NsNm) As StrOpt
Dim Ns$, Nm$
    Ns = Itm.Ns
    Nm = Itm.Nm
Dim J%
For J = 0 To WyUB
    With Wy(J)
        If .Ns <> Ns Then GoTo Nxt
        If .Nm <> Nm Then GoTo Nxt
        If Not .Done Then Exit Function
        Wrk_LookupVal = StrOpt(.Str)
        Exit Function
    End With
Nxt:
Next
End Function
Friend Sub Exp_Prm()
Dim J%, M As WrkDr
For J = 0 To WyUB
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
Friend Sub Exp_Prm__Tst()
With ZZ
    .Wrk_Brw
    .Exp_Prm
    .Wrk_Brw
End With
End Sub
Friend Function Macro_LookupVal(MacroTerm$) As StrOpt
Macro_LookupVal = Wrk_LookupVal(NsNm_ByMacroTerm(MacroTerm))
End Function
Friend Function Macro_TakOneTerm$(MacroStr$)
Macro_TakOneTerm = TakBet(MacroStr, "{", "}", NoTrim:=True, InclMarker:=True)
End Function
Friend Function Macro_Rpl(MacroStr$) As StrOpt
'MacroStr is a Tp string containing {xxx.xxx}, it is required to replace {xxxx.xxxx} by value lookup from Wy
Dim O$
    O = MacroStr
    Dim M$: M = Macro_TakOneTerm(O)
    While M <> ""
        With Macro_LookupVal(M)
            If Not .Som Then Exit Function
            O = Replace(O, M, .Str)
        End With
        M = Macro_TakOneTerm(O)
    Wend
Macro_Rpl = StrOpt(O)
End Function
Sub Exp_FixWh()
Dim J%, M As WrkDr
For J = 0 To WyUB
    M = Wy(J)
    With M
        If .L3.Op <> eFixWh Then GoTo Nxt
        If .Done Then Stop
        With Macro_Rpl(.L3.Prm)
            If Not .Som Then GoTo Nxt
            M.Done = True
            M.Str = .Str
            Wy(J) = M
        End With
    End With
Nxt:
Next
End Sub
Friend Sub Exp_FixWh__Tst()
With ZZ
    .Exp_Prm
    .Exp_Switch
    '.Wrk_Brw
    .Exp_FixWh
    .Wrk_Brw
End With
End Sub
Friend Sub Exp_ThoseWith__Tst()
With ZZ
    .Exp_Prm
    .Exp_Switch
    .Exp_OffSwitchVal
    .Exp_FixStr
    .Wrk_Brw
    Stop
    .Exp_ThoseWithExp
    .Wrk_Brw
End With
End Sub
Friend Function Switch_LookupVal(Switch$) As StrOpt
Switch_LookupVal = Wrk_LookupVal(NsNm("?", Switch))
End Function

Friend Function Switch_EqNe(Prm$, IsEq As Boolean) As BoolOpt
Dim T1$, T2$
    With Brk(Prm, " ")
        T1 = .S1
        T2 = .S2
    End With
Dim V1$
    Dim V1Opt As StrOpt
    V1Opt = Macro_LookupVal(T1)
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
Switch_EqNe = BoolOpt(Bool)
End Function

Function Switch_Val(Switch$) As BoolOpt
Dim J%
For J = 0 To WyUB
    With Wy(J)
        If .Ns <> "?" Then GoTo Nxt
        If .Nm <> Switch Then GoTo Nxt
        If Not .Done Then Exit Function
        Switch_Val = BoolOpt(.Str = "1")
        Exit Function
    End With
Nxt:
Next
End Function

Friend Function Switch_OffDic() As Dictionary
'Return Dic with Switch value is off
Dim J%, O As New Dictionary
For J = 0 To WyUB
    With Wy(J)
        If .Ns <> "?" Then GoTo Nxt
        If Not .Done Then Stop
        If .Str = "1" Then GoTo Nxt
        O.Add .Nm, Empty
    End With
Nxt:
Next
Set Switch_OffDic = O
End Function

Friend Sub Exp_OffSwitchVal__Tst()
With ZZ
    .Exp_Prm
    .Exp_Switch
    .Wrk_Brw
    .Exp_OffSwitchVal
    .Wrk_Brw
End With
End Sub

Friend Function Switch_DefDic() As Dictionary
Dim J%, O As New Dictionary
For J = 0 To WyUB
    With Wy(J)
        If .Ns <> "?" Then GoTo Nxt
        O.Add .Nm, 1
    End With
Nxt:
Next
Set Switch_DefDic = O
End Function

Friend Function Er_SwitchNotExist() As Variant()
Dim J%, O()
Dim Dic As Dictionary: Set Dic = Switch_DefDic
For J = 0 To WyUB
    With Wy(J).L3
        If .Switch = "" Then GoTo Nxt
        If Dic.Exists(.Switch) Then GoTo Nxt
        Push O, Array(Wy(J).LinI, FmtQQ("Switch[?] not exist", .Switch))
    End With
Nxt:
Next
Er_SwitchNotExist = O
End Function

Friend Sub Exp_OffSwitchVal()
'Any Nm with Switch and the Switch has off value, set them to be done and SwitchVal="0"
If Not Switch_IsDone Then MsgBox ".Exp_OffSwitchVal is called only after Fun-.Exp_Switch.  Calling Fun-Exp-Switch will make Switch_IsDone": Stop
Dim Dic As Dictionary
Set Dic = Switch_OffDic

Dim J%, M As WrkDr
For J = 0 To WyUB
    M = Wy(J)
    With M.L3
        If .Switch = "" Then GoTo Nxt
        If Not Dic.Exists(.Switch) Then GoTo Nxt
        M.Done = True
        M.L3.SwitchVal = "0"
        Wy(J) = M
    End With
Nxt:
Next
End Sub

Friend Function Switch_TermVal(Term$) As BoolOpt
If FstChr(Term) = "?" Then
    Switch_TermVal = Switch_Val(RmvFstChr(Term))
    Exit Function
End If
If FstChr(Term) = "{" And LasChr(Term) = "}" Then
    With Macro_LookupVal(Term)
        If .Som Then Switch_TermVal = BoolOpt(.Str = "1")
    End With
    Exit Function
End If
Stop
End Function

Friend Function Switch_AndOr(Prm$, IsAnd As Boolean) As BoolOpt
Dim TermAy$(): TermAy = SplitSpc(Prm)
Dim ValAy() As Boolean
ReDim ValAy(UB(TermAy))
Dim J%
For J = 0 To UB(TermAy)
    With Switch_TermVal(TermAy(J))
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
Switch_AndOr = BoolOpt(Bool)
End Function
Friend Sub Exp_Switch()
Dim J%
While Exp_Switch_OneCycle
    J = J + 1
    If J > 20 Then Stop
Wend
If Not Switch_IsDone Then Stop 'Remark: all the Switch should be able to evalulate, otherwise stop
End Sub

Friend Function Exp_Switch_OneCycle() As Boolean
'Update all Switches (Ns=?) for .Str and .Done
'Remark: all the Switch should be able to evalulate, otherwise stop
Dim OMore As Boolean    'If any switch cannot be evaluated, set OMore = true and to be return
    Dim J%, M As WrkDr, V1$, V2$
    For J = 0 To WyUB
        M = Wy(J)
        With M
            If .Ns <> "?" Then GoTo Nxt
            If .Done Then GoTo Nxt
            Select Case .L3.Op
            Case eEq, eNe
                With Switch_EqNe(.L3.Prm, .L3.Op = eEq)
                    If .Som Then
                        M.Done = True
                        M.Str = IIf(.Bool, "1", "0")
                        Wy(J) = M          '<====================
                    Else
                        OMore = True ' MOre to expand
                    End If
                End With
            Case eFixAnd, eFixOr
                With Switch_AndOr(.L3.Prm, .L3.Op = eEq)
                    If .Som Then
                        M.Done = True
                        M.Str = IIf(.Bool, "1", "0")
                        Wy(J) = M            '<====================
                    Else
                        OMore = True ' MOre to expand
                    End If
                End With
            Case Else: Stop
            End Select
        End With
Nxt:
    Next
Exp_Switch_OneCycle = OMore
End Function

Friend Function Lin_IsL1(L) As Boolean
Dim C$
C = FstChr(L)
If IsLetter(C) Then Lin_IsL1 = True: Exit Function
Select Case C
Case "?", ">", "#" ' Switch, Input, Expression
Lin_IsL1 = True
End Select
End Function

Friend Function Lin_IsL2(L) As Boolean
If Left(L, 4) = Space(4) Then
    Dim C$: C = Mid(L, 5, 1)
    Lin_IsL2 = IsLetter(C) Or C = "?"
End If
End Function

Friend Function Lin_IsL3(L) As Boolean
If Left(L, 8) = Space(8) Then
    Lin_IsL3 = Mid(L, 9, 1) <> " "
End If
End Function

Friend Function Lin_Lvl(L) As Byte
If Lin_IsL1(L) Then Lin_Lvl = 1: Exit Function
If Lin_IsL2(L) Then Lin_Lvl = 2: Exit Function
If Lin_IsL3(L) Then Lin_Lvl = 3: Exit Function
Lin_Lvl = 99
End Function

Friend Function Op(OpStr$) As eOp
Dim O As eOp
Select Case OpStr
Case "$": O = eMac
Case ".": O = eFixStr
Case ".": O = eFixStr
Case ".And": O = eFixAnd
Case ".Bet": O = eBet
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

Friend Function Op_Chk$(OpStr$)

End Function

Friend Function Op_IsVdt(OpStr$) As Boolean

End Function

Friend Function Op_Sy() As String()
Dim O$(), J&
For J = 0 To MdByNm("bb_Lib_Sql3").EnmMbrCnt("eOp") - 1
    Push O, OpStr(J)
Next
Op_Sy = O
End Function

Friend Function OpStr(A As eOp)
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

Friend Function Wrk_Dic() As Dictionary
Dim J%
Dim O As New Dictionary
Exp
For J = 0 To WyUB
    With Wy(J)
        O.Add .Ns & "." & .Nm, .Str
    End With
Next
Set Wrk_Dic = O
End Function

Friend Function Sql3_Dry() As Variant()
'Dr = Ns Nm Str
Dim Dic As Dictionary
Set Dic = Wrk_Dic
Dim ODry(), Dr()
Dim Wrk_ As WrkDr
Dim Nm$, Ns$, Str$, J%
For J = 0 To WyUB
    Wrk_ = Wy(J)
    With Wrk_
        If Not .Done Then Stop
        Dr = Array(.Ns, .Nm, .Str)
        Push ODry, Dr
    End With
Next
Sql3_Dry = ODry
End Function
Friend Function Sql3_IsEmpty() As Boolean
Sql3_IsEmpty = AyIsEmpty(A_)
End Function
Friend Sub Sql3_Brw()
If Sql3_IsEmpty Then Exit Sub
Dim Dry As New Dry
    Dim Dr
    Dim L
    For Each L In A_
        Dry.AddDr Array(Lin_Lvl(L), L)
    Next
Dim Fny$()
    Fny = SplitSpc("Lvl Str")
Drs(Fny, Dry).Brw
End Sub

Friend Function Ly_No2Dash() As String()
If AyIsEmpty(A_) Then Exit Function
Dim O$(), I
For Each I In Ly
    Push O, Brk1(I, "--", NoTrim:=True).S1
Next
Ly_No2Dash = O
End Function

Friend Function Ly_No3Dash() As String()
Dim O$(), I
For Each I In A_
    Push O, Brk1(I, "---", NoTrim:=True).S1
Next
Ly_No3Dash = O
End Function
Friend Function Ly_Validated() As String()
Dim O$(): O = Ly_No3Dash
If IsEr Then Exit Function
Dim I&, Dr
For Each Dr In ErDry
    I = Dr(0)
    O(I) = O(I) & " --- " & Dr(1)
Next
Dim Er$: Er = FmtQQ("--- [?] error(s)", Sz(ErDry))
Ly_Validated = Ay(O).RmvEmptyEleAtEnd.Push(Er).Ay_
End Function

Friend Function Wrk_Drs() As Drs
Dim Dry As New Dry
    Dim J%
    For J = 0 To WyUB
        With Wy(J)
            Dry.AddDr Array(.LinI, .Ns, IIf(.Done, "*", ""), .Nm, .L3.Switch, .L3.SwitchVal, OpStr(.L3.Op), .L3.Prm, .L3.L3, .Str)
        End With
    Next
Dim Fny$(): Fny = SplitSpc("LinI Ns Done Nm Switch SwitchVal Op Prm L3 Str")
Wrk_Drs = Drs(Fny, Dry)
End Function


Friend Function Wrk_WyIsEmpty(A() As WrkDr) As Boolean
End Function
Friend Property Get WySz&()
On Error Resume Next
WySz = UBound(Wy) + 1
End Property
Friend Property Get WyUB&()
WyUB = WySz - 1
End Property

Friend Property Get WyIsEmpty() As Boolean
WyIsEmpty = WySz = 0
End Property

Friend Function WyIsDone() As Boolean
Dim J%
For J = 0 To WyUB
    If Not Wy(J).Done Then Exit Function
Next
WyIsDone = True
End Function

Friend Sub WyPush(M As WrkDr)
Dim N%: N = WySz
ReDim Preserve A_(N)
Wy(N) = M
End Sub

Friend Sub WyPushAy(Ay() As WrkDr)
Dim J%
For J = 0 To WySz - 1
    WyPush Ay(J)
Next
End Sub


Friend Sub ZZSql3_FtFix()
Dim O$(): O = ZZSql3_Ly
Dim J%
For J = 0 To UB(O)
    O(J) = Replace(O(J), Chr(160), " ")
Next
'AyWrt O, ZZSql3_Ft
End Sub

Friend Function ZZSql3_Ly() As String()
'ZZSql3_Ly = FtLy(ZZSql3_Ft)
End Function

Friend Sub Op_Sy__Tst()
'AyDmp Op_Sy
End Sub

Friend Sub Sql3_Brw__Tst()
ZZ.Sql3_Brw
End Sub

Friend Sub Ly_Validated__Tst()
'AyBrw ZZ.Ly_Validated
End Sub
Private Sub Sql3_WrtEr__Tst()
With ZZ
If .Sql3_WrtEr(ZZSql3_Ft) Then ZZEdt
End With
End Sub

Friend Function Sql3_WrtEr(Ft) As Boolean
Dim Ly$(): Ly = Ly_Validated: If AyIsEmpty(Ly) Then Exit Function
'If AyIsEq(Ly, A_) Then Exit Function
'AyWrt Ly, Ft
Sql3_WrtEr = True
End Function
Friend Sub Wrk_Brw()
'Wrk_Drs.Ws Vis:=True
End Sub
Private Sub Wrk_Wy__Tst()
ZZ.Wrk_Brw
End Sub

Sub Tst()
Er_Dry__Tst
'Exp__Tst
'Exp_FixFm__Tst
'Exp_FixStr__Tst
'Exp_FixUpd__Tst
'Exp_FixWh__Tst
'Exp_OffSwitchVal__Tst
'Exp_Prm__Tst
'Exp_Switch__Tst
'Exp_ThoseWith__Tst
'Wrk_Wy__Tst
'Ly_Validated__Tst
'Op_Sy__Tst
'Sql3_Brw__Tst
'Sql3_WrtEr__Tst
End Sub
