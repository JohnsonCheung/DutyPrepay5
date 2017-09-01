Attribute VB_Name = "Sql3"
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
    eMacWh     ' [$Wh] means <Prm> is a Macro String to be used in Sql-Where
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
    OpStr As String
    Op As eOp
    Prm As String    ' RestTerm of L3
End Type
Private Type WrkDr
    Ns As String
    Nm As String
    L3 As L3
    LinI As Integer
End Type
Private Type WrkDrOpt
    Som As Boolean
    WrkDr As WrkDr
End Type

Private Type ExpPrm
    PrmDic As Dictionary
    RestWy() As WrkDr
End Type
Private Type ExpSwitch
    SwitchDic As Dictionary
    RestWy() As WrkDr
End Type
Private Type ExpFixStr
    FixStrDic As Dictionary
    RestWy() As WrkDr
End Type
Private Type ExpFixFm
    FmDic As Dictionary
    RestWy() As WrkDr
End Type
Private Type ExpFixInto
    IntoDic As Dictionary
    RestWy() As WrkDr
End Type
Private Type ExpFixWh
    WhDic As Dictionary
    RestWy() As WrkDr
End Type
Private Type ExpFixUpd
    UpdDic As Dictionary
    RestWy() As WrkDr
End Type
Private Type ExpFixLeftJn
    LeftJnDic As Dictionary
    RestWy() As WrkDr
End Type
Private Type ExpFixJn
    JnDic As Dictionary
    RestWy() As WrkDr
End Type
Private Type ExpFixSelDis
    SelDisDic As Dictionary
    RestWy() As WrkDr
End Type
Private Type ExpFixDrp
    DrpDic As Dictionary
    RestWy() As WrkDr
End Type
Private Type ExpExpTerm
    ExpTermDic As Dictionary
    RestWy() As WrkDr
End Type

Sub AA()
Sql3_Wy__Pass1__Tst
End Sub

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

Function Exp_Prm(Wy() As WrkDr) As ExpPrm
Dim IdxAy%(): IdxAy = Wy_IdxAy_Ns(Wy, "Prm")
Dim ODic As New Dictionary
    Dim J%
    Dim K$, V$
    For J = 0 To UB(IdxAy)
        With Wy(IdxAy(J))
            V = .L3.Prm
            K = .Ns & "." & .Nm
            ODic.Add K, V
        End With
    Next
Set Exp_Prm.PrmDic = ODic
Exp_Prm.RestWy = Wy_RmvItms(Wy, IdxAy)
End Function

Sub Main()
Sql3ExpandedDrs__Tst
End Sub

Sub Sql3_Edt()
FtBrw ZZSql3_Ft
End Sub

Function Sql3ExpandedDrs(Sql3Ly$(), PrmLy$()) As Drs
Dim Wy() As WrkDr:
Wy = Sql3_Wy(Sql3Ly)
Exp Wy
Dim O As Drs
    O.Fny = SplitSpc("Ns Nm Str")
    O.Dry = Wy_Dry(Wy)
Sql3ExpandedDrs = O
End Function

Function Switch_AndOr(PrmDic As Dictionary, SwitchDic As Dictionary, Prm$, IsAnd As Boolean) As BoolOpt
Dim TermAy$(): TermAy = SplitSpc(Prm)
Dim ValAy() As Boolean
    ReDim ValAy(UB(TermAy))
    Dim J%
    For J = 0 To UB(TermAy)
        With Switch_TermVal(PrmDic, SwitchDic, TermAy(J))
            If Not .Som Then
                Exit Function
            End If
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

Function Switch_Val(Wy() As WrkDr, Switch$) As BoolOpt
Dim J%
For J = 0 To Wy_UB(Wy)
    With Wy(J)
        If .Ns <> "?" Then GoTo Nxt
        If .Nm <> Switch Then GoTo Nxt
        'If Not .Done Then Exit Function
        Stop
'        Switch_Val = SomBool(.Str = "1")
        Exit Function
    End With
Nxt:
Next
End Function

Function Wy_RmvItms(A() As WrkDr, IdxAy%()) As WrkDr()
If AyIsEmpty(IdxAy) Then Wy_RmvItms = A: Exit Function
Dim O() As WrkDr, J%
For J = 0 To Wy_UB(A)
    If Not AyHas(IdxAy, J) Then Wy_Push O, A(J)
Next
Wy_RmvItms = O
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
        If .Op = eOp.eUnknown Then Push O, Array(Wy(J).LinI, FmtQQ("Invalid Op[?]", .OpStr))
    End With
Next
Er_InvalidOp = O
End Function

Private Function Er_NoPrm(Wy() As WrkDr) As Variant()
Dim J%
For J = 0 To Wy_UB(Wy)
    If Wy(J).Ns = "Prm" Then Exit Function
Next
Er_NoPrm = Array(Array(0, "Warning: No Prml namespace"))
End Function

Private Function Er_NoSql(Wy() As WrkDr) As Variant()
Dim J%
For J = 0 To Wy_UB(Wy)
    If Wy(J).Ns = "Sql" Then Exit Function
Next
Er_NoSql = Array(Array(0, "Warning: No Sql namespace"))
End Function

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

Private Function Er_UpdMstHavNamWithPondSign(Wy() As WrkDr) As Variant()

End Function

Private Sub Exp(Wy() As WrkDr) 'Expanding Wy
Dim A As ExpPrm: A = Exp_Prm(Wy)
Dim B As ExpSwitch: B = Exp_Switch(A)
Dim C As ExpFixStr: C = Exp_FixStr(B)
Dim Dic As Dictionary
Exp_FixFm Wy
Exp_FixInto Wy
Exp_FixUpd Wy
Exp_FixWh Wy
Exp_ThoseWithExp Dic, Wy
End Sub

Private Function Exp_ExpTerm(Wy() As WrkDr, SwitchDic As Dictionary) As ExpExpTerm
Dim IdxAy%(): IdxAy = Wy_IdxAy_Op(Wy, eOp.eExpTerm)
Dim ODic As New Dictionary
    Dim J%, M As WrkDr
    For J = 0 To UB(IdxAy)
        M = Wy(IdxAy(J))
        Dim K$, V$
        K = M.Ns & "." & M.Nm
        V = Exp_ExpTerm__Val(M.L3.Prm)
        ODic.Add K, V
    Next
Set Exp_ExpTerm.ExpTermDic = ODic
Exp_ExpTerm.RestWy = Wy_RmvItms(Wy, IdxAy)
End Function

Private Function Exp_ExpTerm__Val$(Prm$)
Exp_ExpTerm__Val$ = Prm
End Function

Private Function Exp_FixDrp(Wy() As WrkDr) As ExpFixDrp
Dim IdxAy%(): IdxAy = Wy_IdxAy_Op(Wy, eOp.eFixDrp)
Dim ODic As New Dictionary
    Dim J%, M As WrkDr
    For J = 0 To UB(IdxAy)
        M = Wy(IdxAy(J))
        Dim K$, V$
        K = M.Ns & "." & M.Nm
        V = Exp_FixDrp__Sqs(M.L3.Prm)
        ODic.Add K, V
    Next
Set Exp_FixDrp.DrpDic = ODic
Exp_FixDrp.RestWy = Wy_RmvItms(Wy, IdxAy)
End Function

Private Function Exp_FixDrp__Sqs$(TblNmLvs$)
Exp_FixDrp__Sqs$ = TblNmLvs
End Function

Private Function Exp_FixFm(Wy() As WrkDr) As ExpFixFm
Dim IdxAy%(): IdxAy = Wy_IdxAy_Op(Wy, eFixFm)
Dim ODic As New Dictionary
    Dim J%
    Dim M As WrkDr
    For J = 0 To UB(IdxAy)
        M = Wy(IdxAy(J))
        If Trim(M.L3.Prm) = "" Then Stop
        Dim K$: K = M.Ns & "." & M.Nm
        Dim V$: V = "|  From " & M.L3.Prm
        ODic.Add K, V
    Next
Exp_FixFm.RestWy = Wy_RmvItms(Wy, IdxAy)
Set Exp_FixFm.FmDic = ODic
End Function

Private Function Exp_FixInto(Wy() As WrkDr) As ExpFixInto
Dim IdxAy%(): IdxAy = Wy_IdxAy_Op(Wy, eFixInto)
Dim ODic As New Dictionary
    Dim J%, M As WrkDr
    For J = 0 To UB(IdxAy)
        M = Wy(IdxAy(J))
        Dim K$: K = M.Ns & "." & M.Nm
        Dim V$: V = "|  Into #" & RmvPfx(M.Nm, "?")
        ODic.Add K, V
    Next
Set Exp_FixInto.IntoDic = ODic
Exp_FixInto.RestWy = Wy_RmvItms(Wy, IdxAy)
End Function

Private Function Exp_FixJn(Wy() As WrkDr) As ExpFixJn
Dim IdxAy%(): IdxAy = Wy_IdxAy_Op(Wy, eOp.eFixJn)
Dim ODic As New Dictionary
    Dim J%, M As WrkDr
    For J = 0 To UB(IdxAy)
        M = Wy(IdxAy(J))
        Dim K$, V$
        K = M.Ns & "." & M.Nm
        V = "|  Inner Join " & M.L3.Prm
        ODic.Add K, V
    Next
Set Exp_FixJn.JnDic = ODic
Exp_FixJn.RestWy = Wy_RmvItms(Wy, IdxAy)
End Function

Private Function Exp_FixLeftJn(Wy() As WrkDr) As ExpFixLeftJn
Dim IdxAy%(): IdxAy = Wy_IdxAy_Op(Wy, eOp.eFixLeftJn)
Dim ODic As New Dictionary
    Dim J%, M As WrkDr
    For J = 0 To UB(IdxAy)
        M = Wy(IdxAy(J))
        Dim K$, V$
        K = M.Ns & "." & M.Nm
        V = "|  Left Join " & M.L3.Prm
        ODic.Add K, V
    Next
Set Exp_FixLeftJn.LeftJnDic = ODic
Exp_FixLeftJn.RestWy = Wy_RmvItms(Wy, IdxAy)
End Function

Private Function Exp_FixSelDis(Wy() As WrkDr) As ExpFixSelDis
Dim IdxAy%(): IdxAy = Wy_IdxAy_Op(Wy, eOp.eFixSelDis)
Dim ODic As New Dictionary
    Dim J%, M As WrkDr
    For J = 0 To UB(IdxAy)
        M = Wy(IdxAy(J))
        Dim K$, V$
        K = M.Ns & "." & M.Nm
        V = "Select Distinct " & M.L3.Prm
        ODic.Add K, V
    Next
Set Exp_FixSelDis.SelDisDic = ODic
Exp_FixSelDis.RestWy = Wy_RmvItms(Wy, IdxAy)
End Function

Private Function Exp_FixStr(A As ExpSwitch) As ExpFixStr
Dim SwitchDic As Dictionary:    Set SwitchDic = A.SwitchDic
Dim Wy() As WrkDr:              Wy = A.RestWy
Dim IdxAy%():                   IdxAy = Wy_IdxAy_Op(Wy, eFixStr)
Dim ODic As New Dictionary
    Dim J%
    Dim M As WrkDr
    For J = 0 To UB(IdxAy)
        M = Wy(IdxAy(J))
        Dim K$: K = M.Ns & "." & M.Nm
        Dim V$
            If M.L3.Switch = "" Then
                V = M.L3.Prm
            Else
                Dim S$: S = "?" & M.L3.Switch
                If Not SwitchDic.Exists(S) Then Stop
                If SwitchDic(S) = "1" Then
                    V = M.L3.Prm
                Else
                    V = ""
                End If
            End If
        ODic.Add K, V
    Next
Set Exp_FixStr.FixStrDic = ODic
Exp_FixStr.RestWy = Wy_RmvItms(Wy, IdxAy)
End Function

Private Function Exp_FixUpd(Wy() As WrkDr) As ExpFixUpd
Dim IdxAy%(): IdxAy = Wy_IdxAy_Op(Wy, eOp.eFixUpd)
Dim ODic As New Dictionary
    Dim J%, M As WrkDr
    For J = 0 To UB(IdxAy)
        M = Wy(IdxAy(J))
        Dim K$, V$
        K = M.Ns & "." & M.Nm
        V = "Update #" & Brk(M.Nm, "#").S1
        ODic.Add K, V
    Next
Set Exp_FixUpd.UpdDic = ODic
Exp_FixUpd.RestWy = Wy_RmvItms(Wy, IdxAy)
End Function

Private Function Exp_FixWh(Wy() As WrkDr) As ExpFixWh
Dim IdxAy%(): IdxAy = Wy_IdxAy_Op(Wy, eFixWh)
Dim ODic As New Dictionary
    Dim J%, M As WrkDr
    For J = 0 To UB(IdxAy)
        M = Wy(IdxAy(J))
        Dim K$: K = M.Ns & "." & M.Nm
        Dim V$: V = "|  Where " & M.L3.Prm
        ODic.Add K, V
    Next
Set Exp_FixWh.WhDic = ODic
Exp_FixWh.RestWy = Wy_RmvItms(Wy, IdxAy)
End Function

Private Function Exp_Switch(A As ExpPrm) As ExpSwitch
Dim Wy() As WrkDr:  Wy = A.RestWy
Dim IdxAy%():       IdxAy = Wy_IdxAy_Ns(Wy, "?")
Dim ODic As New Dictionary
    Dim J%
    Dim V$, K$
    Dim B As BoolOpt
    For J = 0 To UB(IdxAy)
        With Wy(IdxAy(J))
            Select Case .L3.Op
            Case eEq, eNe
                B = Switch_EqNe(A.PrmDic, ODic, .L3.Prm, .L3.Op = eEq)
                If B.Som Then
                    V = IIf(B.Bool, "1", "0")
                    K = .Ns & .Nm
                    ODic.Add K, V
                End If
            Case eFixAnd, eFixOr
                B = Switch_AndOr(A.PrmDic, ODic, .L3.Prm, .L3.Op = eEq)
                If B.Som Then
                    V = IIf(B.Bool, "1", "0")
                    K = .Ns & .Nm
                    ODic.Add K, V
                End If
            Case Else: Stop
            End Select
        End With
    Next
Set Exp_Switch.SwitchDic = ODic
Exp_Switch.RestWy = Wy_RmvItms(Wy, IdxAy)
End Function

Private Function Exp_ThoseWithExp(Dic As Dictionary, OWy() As WrkDr) As Boolean
'Return true if all done
Dim J%, M As WrkDr
For J = 0 To UBound(OWy)
    M = OWy(J)
    With M
        Stop
'        If .Done Then GoTo Nxt
        If Not Op_IsExp(.L3.Op) Then GoTo Nxt
        With Macro_ExpTermLis(Dic, .Ns, .Nm, .L3.Prm)
            If Not .Som Then GoTo Nxt
'            M.Str = Exp_ThoseWithExp_Str(M.L3.Op, .Sy)
'            M.Done = True
            Exp_ThoseWithExp = False
        End With
    End With
Nxt:
Next
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

Private Function L3_Brk(L3$) As L3
Dim L$: L = Trim(L3)
If L3 = "" Then Exit Function
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
End Function

Private Function Lin_WrkDr(Sql3Lin, LinI%) As WrkDr
Dim A$: A = Sql3Lin
Dim O As WrkDr
With O
    .LinI = LinI%
    Select Case Lin_Lvl(Sql3Lin)
    Case 1
        .Ns = ParseTerm(A)
        .Nm = ParseTerm(A)
        .L3 = L3_Brk(A)
    Case 2
        .Nm = ParseTerm(A)
        .L3 = L3_Brk(A)
    Case 3
        .L3 = L3_Brk(Trim(A))
    Case Else: Stop
    End Select
End With
Lin_WrkDr = O
End Function

Private Function Macro_ExpTermLis(Dic As Dictionary, Ns$, Nm$, TermLis) As SyOpt
'Each terms in {TermLis} is term-list required to be expanded into a str
'Each term, Ns.Nm.Term, will be used to look up from Dic
'Return None is any term cannot be lookup in Dic
Dim Ay$(): Ay = SplitLvs(TermLis)
Dim O$(), T, Pfx$, S$
Pfx = Ns & "." & Nm & "."
For Each T In Ay
    S = Pfx & T
    With DicVal(Dic, S)
        If Not .Som Then Exit Function
        Push O, .V
    End With
Next
Macro_ExpTermLis = SomSy(O)
End Function

Private Function Macro_Rpl(Dic As Dictionary, Wy() As WrkDr, MacroStr$) As StrOpt
Dim O$
    O = MacroStr
    Dim M$: M = TakBet(O, "{", "}", InclMarker:=True)
    While M <> ""
        With DicVal(Dic, M)
            If Not .Som Then Exit Function
            O = Replace(O, M, .V)
        End With
        M = TakBet(O, "{", "}", InclMarker:=True)
    Wend
Macro_Rpl = SomStr(O)
End Function

Private Function Op(OpStr$) As eOp
Dim O As eOp
Select Case OpStr
Case "$": O = eMac
Case "$And": O = eMacAnd
Case "$Wh": O = eMacWh
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

Private Function Op_AlwSwitchOpAy() As eOp()
Dim O() As eOp, I
For Each I In Array(eOp.eFixFm, eOp.eFixGp, eOp.eExpGp, eOp.eFixInto, eOp.eFixSelDis, eOp.eExpSelDis, eOp.eMac, eOp.eExpTerm, eOp.eExpComma, _
    eOp.eFixLeftJn, eOp.eFixJn, eOp.eExpJn, eOp.eExpLeftJn, _
    eOp.eExpSel, eOp.eFixSel, eOp.eMacAnd, eOp.eMacOr, eOp.eFixStr)
    Push O, I
Next
Op_AlwSwitchOpAy = O
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

Private Function Op_Chk$(OpStr$)

End Function

Private Function Op_IsAlwSwitch(A As eOp) As Boolean
Op_IsAlwSwitch = AyHas(Op_AlwSwitchOpAy, A)
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

Private Function Op_IsVdt(OpStr$) As Boolean

End Function

Private Function Op_Sy() As String()
Dim O$(), J&
For J = 0 To EnmMbrCnt("eOp", Md("Sql3")) - 1
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
Case eMacWh: O = "$Wh"
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

Private Function SomWrkDr(A As WrkDr) As WrkDrOpt
SomWrkDr.Som = True
SomWrkDr.WrkDr = A
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
Dim ErDry(): ErDry = Er_Dry(Sql3_Wy(No3Dash_Sql3Ly))
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

Private Function Sql3_WrkDrOpt(Sql3Lin, LinI%) As WrkDrOpt
If Lin_Lvl(Sql3Lin) = 0 Then Exit Function
Sql3_WrkDrOpt = SomWrkDr(Lin_WrkDr(Sql3Lin, LinI))
End Function

Private Function Sql3_WrtEr(Ft) As Boolean
Sql3_Rmv3DashInFt Ft
Dim Ly$(): Ly = FtLy(Ft)
Dim Ly1$(): Ly1 = Sql3_ValidatedLy(Ly): If AyIsEmpty(Ly1) Then Exit Function
If AyIsEq(Ly, Ly1) Then Exit Function
AyWrt Ly1, Ft
Sql3_WrtEr = True
End Function

Private Function Sql3_Wy(Sql3_Ly$()) As WrkDr()
Dim A() As WrkDr: A = Sql3_Wy__Pass1(Sql3_Ly)
Dim B() As WrkDr: B = Sql3_Wy__Pass2(A)
Dim O() As WrkDr
    Dim J%, M As WrkDr
    For J = 0 To Wy_UB(B)
        M = B(J)
        If M.L3.L3 <> "" Then Wy_Push O, M
    Next
Sql3_Wy = O
End Function

Private Function Sql3_Wy__Pass1(Sql3_Ly$()) As WrkDr()
Dim O() As WrkDr
    Dim L, LinI%, A As WrkDrOpt
    For Each L In Sql3_Ly
        With Sql3_WrkDrOpt(L, LinI)
            If .Som Then
                Wy_Push O, .WrkDr
            End If
        End With
        LinI = LinI + 1
    Next
Sql3_Wy__Pass1 = O
End Function

Private Function Sql3_Wy__Pass2(Wy() As WrkDr) As WrkDr()
Dim O() As WrkDr
    O = Wy
    Dim J%
    Dim LasNs$, LasNm$
    For J = 0 To Wy_UB(O)
        If O(J).Ns = "" Then
            O(J).Ns = LasNs
            If O(J).Nm = "" Then
                O(J).Nm = LasNm
            End If
        End If
        LasNs = O(J).Ns
        LasNm = O(J).Nm
    Next
Sql3_Wy__Pass2 = O
End Function

Private Function Switch_Dic(Wy() As WrkDr) As Dictionary
'Return Dic with Switch value is off
Dim J%, O As New Dictionary
For J = 0 To Wy_UB(Wy)
    With Wy(J)
        If .Ns <> "?" Then GoTo Nxt
    Stop
'        If Not .Done Then Stop
'        O.Add .Nm, .Str = "1"
    End With
Nxt:
Next
Set Switch_Dic = O
End Function

Private Function Switch_EqNe(PrmDic As Dictionary, SwitchDic As Dictionary, Prm$, IsEq As Boolean) As BoolOpt
Dim T1$, T2$
    With Brk(Prm, " ")
        T1 = .S1
        T2 = .S2
    End With
    If FstChr(T1) = "{" And LasChr(T1) = "}" Then
        T1 = RmvLasChr(RmvFstChr(T1))
    ElseIf FstChr(T1) = "?" Then
    Else
        Stop
    End If
Dim Dic As Dictionary
    Set Dic = DicAdd(PrmDic, SwitchDic)
    
Dim V1$
    Dim V1Opt As VOpt
    V1Opt = DicVal(Dic, T1)
    If Not V1Opt.Som Then Exit Function
    V1 = V1Opt.V
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

Private Function Switch_TermVal(PrmDic As Dictionary, SwitchDic As Dictionary, Term$) As BoolOpt
If FstChr(Term) = "?" Then
    With DicVal(SwitchDic, Term)
        If .Som Then Switch_TermVal = SomBool(.V)
    End With
    Exit Function
End If
If FstChr(Term) = "{" And LasChr(Term) = "}" Then
    Dim A$
    A = RmvLasChr(RmvFstChr(Term))
    With DicVal(PrmDic, A)
        If .Som Then Switch_TermVal = SomBool(.V = "1")
    End With
    Exit Function
End If
Stop
End Function

Private Sub Wy_Brw(Wy() As WrkDr, Optional SwitchDic As Dictionary)
DrsWs Wy_Drs(Wy, SwitchDic)
End Sub

Private Function Wy_Dic(Wy() As WrkDr) As Dictionary
If Wy_IsEmpty(Wy) Then Exit Function
Dim J%, W As WrkDr
Dim U%, K$
    U = UBound(Wy)
Dim O As New Dictionary
    For J = 0 To U
        With Wy(J)
            K = .Ns & "." & .Nm
            Stop
'            O.Add K, .Str
        End With
    Next
Set Wy_Dic = O
End Function

Private Function Wy_Dr(A As WrkDr, SwitchDic As Dictionary) As Variant()
Dim SwitchVal$
    If A.L3.Switch <> "" Then
        With DicVal(SwitchDic, "?" & A.L3.Switch)
            If .Som Then
                SwitchVal = .V
            Else
                SwitchVal = "{?}"
            End If
        End With
    End If
With A
    Wy_Dr = Array(.LinI, .Ns, .Nm, .L3.Switch, SwitchVal, OpStr(.L3.Op), .L3.Prm, .L3.L3)
End With
End Function

Private Function Wy_Drs(Wy() As WrkDr, Optional SwitchDic As Dictionary) As Drs
Dim Dic As Dictionary
    If IsNothing(SwitchDic) Then
        Set Dic = New Dictionary
    Else
        Set Dic = SwitchDic
    End If

Dim ODry()
    Dim J%
    For J = 0 To Wy_UB(Wy)
        Push ODry, Wy_Dr(Wy(J), Dic)
    Next
Dim O As Drs
O.Fny = Wy_Fny
O.Dry = ODry
Wy_Drs = O
End Function

Private Function Wy_Dry(Wy() As WrkDr) As Variant()
Wy_Dry = DicDry(Wy_Dic(Wy))
End Function

Private Function Wy_Fny() As String()
Wy_Fny = SplitSpc("LinI Ns Nm Switch SwitchVal Op Prm L3")
End Function

Private Function Wy_IdxAy_Ns(Wy() As WrkDr, Ns$) As Integer()
Dim O%()
    Dim J%
    For J = 0 To Wy_UB(Wy)
        If Wy(J).Ns = Ns Then Push O, J
    Next
Wy_IdxAy_Ns = O
End Function

Private Function Wy_IdxAy_Op(Wy() As WrkDr, Op As eOp) As Integer()
Dim O%()
    Dim J%
    For J = 0 To Wy_UB(Wy)
        If Wy(J).L3.Op = Op Then Push O, J
    Next
Wy_IdxAy_Op = O
End Function

Private Function Wy_IsEmpty(A() As WrkDr) As Boolean
Wy_IsEmpty = Wy_Sz(A) = 0
End Function

Private Sub Wy_Push(OAy() As WrkDr, M As WrkDr)
Dim N%: N = Wy_Sz(OAy)
ReDim Preserve OAy(N)
OAy(N) = M
End Sub

Private Sub Wy_PushAy(OAy() As WrkDr, Ay() As WrkDr)
Dim J%
For J = 0 To Wy_Sz(Ay) - 1
    Wy_Push OAy, Ay(J)
Next
End Sub

Private Function Wy_Sz&(Wy() As WrkDr)
On Error Resume Next
Wy_Sz = UBound(Wy) + 1
End Function

Private Function Wy_UB&(Wy() As WrkDr)
Wy_UB = Wy_Sz(Wy) - 1
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

Private Function ZZSql3_Wy() As WrkDr()
ZZSql3_Wy = Sql3_Wy(ZZSql3_Ly)
End Function

Private Function ZZWy() As WrkDr()
If Sql3_WrtEr(ZZSql3_Ft) Then FtBrw ZZSql3_Ft: Stop
ZZWy = Sql3_Wy(ZZSql3_Ly)
End Function

Private Sub Exp__Tst()
Sql3_Rmv3DashInFt ZZSql3_Ft
If Sql3_WrtEr(ZZSql3_Ft) Then FtBrw ZZSql3_Ft: Exit Sub
Dim Wy() As WrkDr: Wy = Sql3_Wy(ZZSql3_Ly)
Exp Wy
Wy_Brw Wy
End Sub

Private Sub Exp_ExpTerm__Tst()
Dim Wy() As WrkDr: Wy = ZZWy
Dim A As ExpPrm: A = Exp_Prm(Wy)
Dim B As ExpSwitch: B = Exp_Switch(A)
Dim C As ExpFixStr: C = Exp_FixStr(B)
Dim D As ExpFixFm: D = Exp_FixFm(C.RestWy)
Dim E As ExpFixInto: E = Exp_FixInto(D.RestWy)
Dim F As ExpFixWh: F = Exp_FixWh(E.RestWy)
Dim G As ExpFixUpd: G = Exp_FixUpd(F.RestWy)
Dim H As ExpFixLeftJn: H = Exp_FixLeftJn(G.RestWy)
Dim I As ExpFixJn: I = Exp_FixJn(H.RestWy)
Dim K As ExpFixSelDis: K = Exp_FixSelDis(I.RestWy)
Dim L As ExpFixDrp: L = Exp_FixDrp(K.RestWy)
Dim M As ExpExpTerm: M = Exp_ExpTerm(L.RestWy, B.SwitchDic)
DicBrw M.ExpTermDic
Dim Dif%
    Dim Bef%: Bef = Wy_UB(L.RestWy)
    Dim Aft%: Aft = Wy_UB(M.RestWy)
    Dif = Bef - Aft
Debug.Assert M.ExpTermDic.Count = Dif
Wy_Brw M.RestWy, B.SwitchDic
End Sub

Private Sub Exp_FixDrp__Tst()
Dim Wy() As WrkDr: Wy = ZZWy
Dim A As ExpPrm: A = Exp_Prm(Wy)
Dim B As ExpSwitch: B = Exp_Switch(A)
Dim C As ExpFixStr: C = Exp_FixStr(B)
Dim D As ExpFixFm: D = Exp_FixFm(C.RestWy)
Dim E As ExpFixInto: E = Exp_FixInto(D.RestWy)
Dim F As ExpFixWh: F = Exp_FixWh(E.RestWy)
Dim G As ExpFixUpd: G = Exp_FixUpd(F.RestWy)
Dim H As ExpFixLeftJn: H = Exp_FixLeftJn(G.RestWy)
Dim I As ExpFixJn: I = Exp_FixJn(H.RestWy)
Dim K As ExpFixSelDis: K = Exp_FixSelDis(I.RestWy)
Dim L As ExpFixDrp: L = Exp_FixDrp(K.RestWy)
DicBrw L.DrpDic
Dim Dif%
    Dim Bef%: Bef = Wy_UB(K.RestWy)
    Dim Aft%: Aft = Wy_UB(L.RestWy)
    Dif = Bef - Aft
Debug.Assert L.DrpDic.Count = Dif
Wy_Brw L.RestWy, B.SwitchDic
End Sub

Private Sub Exp_FixFm__Tst()
Dim Wy() As WrkDr: Wy = ZZWy
Dim A As ExpPrm: A = Exp_Prm(Wy)
Dim B As ExpSwitch: B = Exp_Switch(A)
Dim C As ExpFixStr: C = Exp_FixStr(B)
Dim D As ExpFixFm: D = Exp_FixFm(C.RestWy)
DicBrw D.FmDic
Wy_Brw D.RestWy
Dim Dif%
    Dim Aft%: Aft = Wy_UB(D.RestWy)
    Dim Bef%: Bef = Wy_UB(C.RestWy)
    Dif = Bef - Aft
Debug.Assert Dif = D.FmDic.Count
End Sub

Private Sub Exp_FixInto__Tst()
Dim Wy() As WrkDr: Wy = ZZWy
Dim A As ExpPrm: A = Exp_Prm(Wy)
Dim B As ExpSwitch: B = Exp_Switch(A)
Dim C As ExpFixStr: C = Exp_FixStr(B)
Dim D As ExpFixFm: D = Exp_FixFm(C.RestWy)
Dim E As ExpFixInto: E = Exp_FixInto(D.RestWy)

DicBrw E.IntoDic
Wy_Brw E.RestWy, B.SwitchDic
Dim Dif%
    Dim Aft%: Aft = Wy_UB(E.RestWy)
    Dim Bef%: Bef = Wy_UB(D.RestWy)
    Dif = Bef - Aft
Debug.Assert Dif = E.IntoDic.Count
End Sub

Private Sub Exp_FixJn__Tst()
Dim Wy() As WrkDr: Wy = ZZWy
Dim A As ExpPrm: A = Exp_Prm(Wy)
Dim B As ExpSwitch: B = Exp_Switch(A)
Dim C As ExpFixStr: C = Exp_FixStr(B)
Dim D As ExpFixFm: D = Exp_FixFm(C.RestWy)
Dim E As ExpFixInto: E = Exp_FixInto(D.RestWy)
Dim F As ExpFixWh: F = Exp_FixWh(E.RestWy)
Dim G As ExpFixUpd: G = Exp_FixUpd(F.RestWy)
Dim H As ExpFixLeftJn: H = Exp_FixLeftJn(G.RestWy)
Dim I As ExpFixJn: I = Exp_FixJn(H.RestWy)
DicBrw I.JnDic
Dim Dif%
    Dim Bef%: Bef = Wy_UB(H.RestWy)
    Dim Aft%: Aft = Wy_UB(I.RestWy)
    Dif = Bef - Aft
Debug.Assert I.JnDic.Count = Dif
Wy_Brw I.RestWy, B.SwitchDic
End Sub

Private Sub Exp_FixLeftJn__Tst()
Dim Wy() As WrkDr: Wy = ZZWy
Dim A As ExpPrm: A = Exp_Prm(Wy)
Dim B As ExpSwitch: B = Exp_Switch(A)
Dim C As ExpFixStr: C = Exp_FixStr(B)
Dim D As ExpFixFm: D = Exp_FixFm(C.RestWy)
Dim E As ExpFixInto: E = Exp_FixInto(D.RestWy)
Dim F As ExpFixWh: F = Exp_FixWh(E.RestWy)
Dim G As ExpFixUpd: G = Exp_FixUpd(F.RestWy)
Dim H As ExpFixLeftJn: H = Exp_FixLeftJn(G.RestWy)
DicBrw H.LeftJnDic
Dim Dif%
    Dim Bef%: Bef = Wy_UB(G.RestWy)
    Dim Aft%: Aft = Wy_UB(H.RestWy)
    Dif = Bef - Aft
Debug.Assert H.LeftJnDic.Count = Dif
Wy_Brw H.RestWy, B.SwitchDic
End Sub

Private Sub Exp_FixSelDis__Tst()
Dim Wy() As WrkDr: Wy = ZZWy
Dim A As ExpPrm: A = Exp_Prm(Wy)
Dim B As ExpSwitch: B = Exp_Switch(A)
Dim C As ExpFixStr: C = Exp_FixStr(B)
Dim D As ExpFixFm: D = Exp_FixFm(C.RestWy)
Dim E As ExpFixInto: E = Exp_FixInto(D.RestWy)
Dim F As ExpFixWh: F = Exp_FixWh(E.RestWy)
Dim G As ExpFixUpd: G = Exp_FixUpd(F.RestWy)
Dim H As ExpFixLeftJn: H = Exp_FixLeftJn(G.RestWy)
Dim I As ExpFixJn: I = Exp_FixJn(H.RestWy)
Dim K As ExpFixSelDis: K = Exp_FixSelDis(I.RestWy)
DicBrw K.SelDisDic
Dim Dif%
    Dim Bef%: Bef = Wy_UB(I.RestWy)
    Dim Aft%: Aft = Wy_UB(K.RestWy)
    Dif = Bef - Aft
Debug.Assert K.SelDisDic.Count = Dif
Wy_Brw K.RestWy, B.SwitchDic
End Sub

Private Sub Exp_FixStr__Tst()
Dim Wy() As WrkDr: Wy = ZZWy
Dim A As ExpPrm: A = Exp_Prm(Wy)
Dim B As ExpSwitch: B = Exp_Switch(A)
Dim C As ExpFixStr: C = Exp_FixStr(B)
DicBrw C.FixStrDic
Dim Dif%
    Dim Bef%: Bef = Wy_UB(B.RestWy)
    Dim Aft%: Aft = Wy_UB(C.RestWy)
    Dif = Bef - Aft
Debug.Assert C.FixStrDic.Count = Dif
Wy_Brw C.RestWy, B.SwitchDic
End Sub

Private Sub Exp_FixUpd__Tst()
Dim Wy() As WrkDr: Wy = ZZWy
Dim A As ExpPrm: A = Exp_Prm(Wy)
Dim B As ExpSwitch: B = Exp_Switch(A)
Dim C As ExpFixStr: C = Exp_FixStr(B)
Dim D As ExpFixFm: D = Exp_FixFm(C.RestWy)
Dim E As ExpFixInto: E = Exp_FixInto(D.RestWy)
Dim F As ExpFixWh: F = Exp_FixWh(E.RestWy)
Dim G As ExpFixUpd: G = Exp_FixUpd(F.RestWy)
DicBrw G.UpdDic
Dim Dif%
    Dim Bef%: Bef = Wy_UB(F.RestWy)
    Dim Aft%: Aft = Wy_UB(G.RestWy)
    Dif = Bef - Aft
Debug.Assert G.UpdDic.Count = Dif
Wy_Brw G.RestWy, B.SwitchDic
End Sub

Private Sub Exp_FixWh__Tst()
Dim Wy() As WrkDr: Wy = ZZWy
Dim A As ExpPrm: A = Exp_Prm(Wy)
Dim B As ExpSwitch: B = Exp_Switch(A)
Dim C As ExpFixStr: C = Exp_FixStr(B)
Dim D As ExpFixFm: D = Exp_FixFm(C.RestWy)
Dim E As ExpFixInto: E = Exp_FixInto(D.RestWy)
Dim F As ExpFixWh: F = Exp_FixWh(E.RestWy)
DicBrw F.WhDic
Dim Dif%
    Dim Bef%: Bef = Wy_UB(E.RestWy)
    Dim Aft%: Aft = Wy_UB(F.RestWy)
    Dif = Bef - Aft
Debug.Assert F.WhDic.Count = Dif
Wy_Brw F.RestWy, B.SwitchDic
End Sub

Private Sub Exp_Prm__Tst()
Dim Wy() As WrkDr: Wy = ZZWy
Dim A As ExpPrm: A = Exp_Prm(Wy) '<== Run
DicBrw A.PrmDic     '<=== Result
Dim Diff%
    Dim Bef%, Aft%
    Bef = Wy_UB(Wy)
    Aft = Wy_UB(A.RestWy)
    Diff = Bef - Aft
Debug.Assert A.PrmDic.Count = Diff
Dim J%
For J = 0 To Wy_UB(A.RestWy)
    With A.RestWy(J)
        Debug.Assert .Ns <> "Prm"
    End With
Next
End Sub

Sub Exp_Switch__Tst()
Dim Wy() As WrkDr: Wy = ZZWy
Dim A As ExpPrm: A = Exp_Prm(Wy)
Dim B As ExpSwitch: B = Exp_Switch(A)
DicBrw B.SwitchDic
Dim Diff%
    Dim Aft%, Bef%
    Bef = Wy_UB(A.RestWy)
    Aft = Wy_UB(B.RestWy)
    Diff = Bef - Aft
Debug.Assert B.SwitchDic.Count = Diff
Wy_Brw B.RestWy
End Sub

Private Sub Exp_ThoseWithExp__Tst()
Dim Wy() As WrkDr: Wy = ZZWy
Dim A As ExpPrm: A = Exp_Prm(Wy)
Dim B As ExpSwitch: B = Exp_Switch(A)
Dim C As ExpFixStr: C = Exp_FixStr(B)
Stop
'Exp_ThoseWithExp Dic, Wy
Wy_Brw Wy
End Sub

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

Private Sub Sql3_Wy__Pass1__Tst()
Dim Ly$(): Ly = ZZSql3_Ly
Dim A() As WrkDr: A = Sql3_Wy__Pass1(Ly)
Wy_Brw A
End Sub

Private Sub Sql3_Wy__Pass2__Tst()
Dim Ly$(): Ly = ZZSql3_Ly
Dim A() As WrkDr: A = Sql3_Wy__Pass1(Ly)
Dim B() As WrkDr: B = Sql3_Wy__Pass2(A)
Wy_Brw A
Wy_Brw B
End Sub

Private Sub Sql3_Wy__Tst()
Sql3_Rmv3DashInFt ZZSql3_Ft
Dim Act() As WrkDr: Act = Sql3_Wy(ZZSql3_Ly)
Wy_Brw Act
Sql3_LyBrw__Tst
End Sub

Sub Sql3ExpandedDrs__Tst()
Dim PrmLy$()
Sql3ExpandedDrs ZZSql3_Ly, PrmLy
End Sub

Private Sub ZZSql3_Ly__Tst()
AyBrw ZZSql3_Ly
End Sub

Sub Tst()
Exp__Tst
Exp_ExpTerm__Tst
Exp_FixDrp__Tst
Exp_FixFm__Tst
Exp_FixInto__Tst
Exp_FixJn__Tst
Exp_FixLeftJn__Tst
Exp_FixSelDis__Tst
Exp_FixStr__Tst
Exp_FixUpd__Tst
Exp_FixWh__Tst
Exp_Prm__Tst
Exp_Switch__Tst
Exp_ThoseWithExp__Tst
Op_Sy__Tst
Sql3_LyBrw__Tst
Sql3_ValidatedLy__Tst
Sql3_WrtEr__Tst
Sql3_Wy__Pass1__Tst
Sql3_Wy__Pass2__Tst
Sql3_Wy__Tst
Sql3ExpandedDrs__Tst
ZZSql3_Ly__Tst
End Sub

