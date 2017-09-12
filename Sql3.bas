Attribute VB_Name = "Sql3"
Option Explicit
Option Compare Database
Public Enum eOp
    eSqlFix '  [#.]
    eExpIn  ' [$In]
    eRun    ' [!]
    eSqlPhrase ' [#]
    'eStr eNbr eFlag eNbrLis eStrLis eFlag are valid only in Ns:Prm
    eFixSql     ' [.Sql]
    eFixEq      ' [.EQ]
    eExpAnd     ' [@And] means <Prm> is term list for "Sql-And"
    eExpComma   ' [@Comma]
    eExpDrp
    eExpGp
    eExpJn
    eExpLeftJn
    eExpOr  ' [@Or] means <Prm> is term list for "Sql-Or"
    eExpSel  ' [@Sel] means <Prm> is term list for "Sql-Select"
    eExpSelDis  ' [@SelDis] means <Prm> is term list for "Sql-Select-Distinct"
    eExp  ' [@]   means <Prm> is term list for sql-statment
    eExpSet   ' [@Set] means <Prm> is term list which will be expanded into "Set <Term> = <Exp-term>, .."
    eExpWh   ' [@Wh] means <Prm> is term list for "Sql-Where"
    eFixAnd  ' [.And] means <Prm> is fixed str for "Sql-And"
    eFixComma ' [.Comma]
    eFixDrp
    eFixFm      ' [.Fm]  means <Prm> is fixed str for "Sql-From"
    eFixGp
    eFixOr  ' [.Or] means <Prm> is fixed str for "Sql-Or"
    eFixSel  ' [.Sel] means <Prm> is fixed str for "Sql-Select"
    eFixSet  ' [.Set] means <Prm> is fixed str for "Sql-Set" to be expanded as Set <Prm>
    eFixSelDis ' [.SelDis] means <Prm> is fixed str for "Sql-Select-Distinct"
    eFixStr    ' [.]   means <Prm> is fixed str
    eFixUpd    ' [.Upd]  means <Prm> is empty "Sql-Update" to be expanded Update #<Nm>
    eFixWh     ' [.Wh] means <Prm> is fixed str for "Sql-Where"
    eFixFlag      ' [.Flag]
    eFixLeftJn ' [.LeftJn] means <Prm> is fixed str for "Sql-Left-Join"
    eFixJn     ' [.Jn] means <Prm> is fixed str for "Sql-inner-Join"
    eMac       ' [$] means <Prm> is a macro string ( a template string with {..} to be expand.  Inside {..} is a <Ns>.<Nm>.
    eMacAnd    ' [$And] means <Prm> is a macro-string
    eMacOr     ' [$Or] means <Prm> is a Macro String to be used in Sql-Or
    eMacWh     ' [$Wh] means <Prm> is a Macro String to be used in Sql-Where
    eFixNbr     ' [.Nbr] means <Prm> is a number
    eFixNbrLis  ' [.NbrLis]
    eFixNe      ' [.NE]
    eFixStrLis  ' [.StrLis]
    eUnknown '
End Enum
Private Type KPD
    K As String   ' Ns.Nm
    P As String   ' L3Prm
    D As Dictionary
End Type
Private Type PrmR
    FunNm As String
    PrmAy() As String
End Type
Private Type L123
    L1 As String
    L2 As String
    L3 As String
    LinI As Integer
End Type
Private Type L123Opt
    Som As Boolean
    L123 As L123
End Type
Private Type L1233
    LinI As Integer
    L1 As String
    L2 As String
    Swtich As String
    OpStr As String
    Prm As String
End Type
Private Type ExpOpt
    Som As Boolean
    L123AyLeft() As L123
    Dic As Dictionary
End Type

Private Type L33
    Switch As String
    OpStr As String
    Prm As String
End Type
Private Type L3
    L3 As String     ' [?<Switch>] <OpTy>[<Op>] [<Prm>]
    Switch As String ' Start with ?, but
    OpStr As String
    Op As eOp
    Prm As String    ' RestTerm of L3
End Type
Private Type Wr
    LinI As Integer
    Ns As String
    Nm As String
    Switch As String
    Op As eOp
    Prm As String
End Type
Private Type WrkDr
    Ns As String
    Nm As String
    L3 As L3
    LinI As Integer
End Type
Private Type Sts
    Wy() As WrkDr
    Dic As Dictionary
End Type
Private Type Sts1
    Wrs As Drs
    Dic As Dictionary
End Type
Private Type StsRun
    RestWy() As WrkDr
    RunDic As Dictionary
End Type

Private Type StsMulNm
    MulNmDic As Dictionary
    RestWy() As WrkDr
End Type

Private Type WrkDrOpt
    Som As Boolean
    WrkDr As WrkDr
End Type
Private Type StsPrm
    PrmDic As Dictionary
    RestWy() As WrkDr
End Type
Private Type StsSwitch
    RestWy() As WrkDr
    SwitchDic As Dictionary
End Type
Private Type StsFixXXX
    RestWy() As WrkDr
    XXXDic As Dictionary
End Type
Private Type StsFixStr
    RestWy() As WrkDr
    StrDic As Dictionary
End Type
Private Type StsFixFm
    RestWy() As WrkDr
    FmDic As Dictionary
End Type
Private Type StsFixWh
    RestWy() As WrkDr
    WhDic As Dictionary
End Type
Private Type StsFixUpd
    RestWy() As WrkDr
    UpdDic As Dictionary
End Type
Private Type StsFixLeftJn
    RestWy() As WrkDr
    LeftJnDic As Dictionary
End Type
Private Type StsFixJn
    RestWy() As WrkDr
    JnDic As Dictionary
End Type
Private Type StsFixSelDis
    RestWy() As WrkDr
    SelDisDic As Dictionary
End Type
Private Type StsFixDrp
    RestWy() As WrkDr
    DrpDic As Dictionary
End Type
Private Type StsExp
    RestWy() As WrkDr
    ExpDic As Dictionary
End Type

Sub AA()
Sql3Ft_Dic__Tst
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

Sub AAA()
L123Ay_L123Pass3Ay__Tst
End Sub

Sub Edt()
ZZSql3Ft_Edt
End Sub

Function L123(LinI%, L1$, L2$, L3$) As L123
Dim O As L123
With O
    .L1 = L1
    .L2 = L2
    .L3 = L3
    .LinI = LinI
End With
L123 = O
End Function

Function Lin_No3DashLin$(Lin)
Lin_No3DashLin = RTrim(RmvAft(Lin, "---"))
End Function

Function Lin_TrmLin$(Lin)
Lin_TrmLin = RTrim(RmvAft(Lin, "--"))
End Function

Sub Main()
Sql3Ft_Dic__Tst
End Sub

Function Sql3Ft_Dic(Sql3Ft$) As Dictionary
Dim Ly$(): Ly = FtLy(Sql3Ft)
Dim L123Ay() As L123
L123Ay = Sql3Ly_L123Ay(Ly)
Set Sql3Ft_Dic = L123Ay_Dic(L123Ay)
End Function

Function Sql3Ly_TrmLy(Sql3Ly$()) As String()
Sql3Ly_TrmLy = AyMapIntoSy(Sql3Ly, "Lin_TrmLin")
End Function

Function SwitchPrm_BoolAy(SwitchPrm$, Dic As Dictionary) As Boolean()
Dim TermAy$()
    TermAy = SplitSpc(SwitchPrm)
Dim O() As Boolean
    ReDim O(UB(TermAy))
    Dim J%
    For J = 0 To UB(TermAy)
        With SwitchTerm_Val(TermAy(J), Dic)
            If Not .Som Then
                Stop
                Exit Function
            End If
            O(J) = .Bool
        End With
    Next
SwitchPrm_BoolAy = O
End Function

Function SwitchPrm_Val_And(SwitchPrm$, Dic As Dictionary) As Boolean
SwitchPrm_Val_And = BoolAy_Or(SwitchPrm_BoolAy(SwitchPrm$, Dic))
End Function

Function SwitchPrm_Val_Eq(SwitchPrm$, Dic As Dictionary) As Boolean
With SwitchPrm_V1V2(SwitchPrm$, Dic)
    SwitchPrm_Val_Eq = .S1 = .S2
End With
End Function

Function SwitchPrm_Val_Ne(SwitchPrm$, Dic As Dictionary) As Boolean
With SwitchPrm_V1V2(SwitchPrm$, Dic)
    SwitchPrm_Val_Ne = .S1 <> .S2
End With
End Function

Function SwitchPrm_Val_Or(SwitchPrm$, Dic As Dictionary) As Boolean
SwitchPrm_Val_Or = BoolAy_Or(SwitchPrm_BoolAy(SwitchPrm$, Dic))
End Function

Function Wy_RmvItms(A() As WrkDr, IdxAy%()) As WrkDr()
If AyIsEmpty(IdxAy) Then Wy_RmvItms = A: Exit Function
Dim O() As WrkDr, J%
For J = 0 To Wy_UB(A)
    If Not AyHas(IdxAy, J) Then Wy_Push O, A(J)
Next
Wy_RmvItms = O
End Function

Function Wy_StsPrm(Wy() As WrkDr) As StsPrm
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
Set Wy_StsPrm.PrmDic = ODic
Wy_StsPrm.RestWy = Wy_RmvItms(Wy, IdxAy)
End Function

Sub ZZSql3Ft_Edt()
FtBrw ZZSql3Ft
End Sub

Private Function Eval_L123(A As L123, Dic As Dictionary) As StrOpt
Dim K$: K = L3_Key(A)
Dim L3Sy$(): L3Sy = SplitCrLf(A.L3)
Dim Vy$()
Dim J%

For J = 0 To UB(L3Sy)
    With Eval_L3Str(L3Sy(J), K, Dic)
        If Not .Som Then Exit Function
        PushNonEmpty Vy, .Str
    End With
Next
Eval_L123 = SomStr(JnVBar(Vy))
End Function

Private Function Eval_L123Ay(A() As L123, Dic As Dictionary) As ExpOpt
Dim O As ExpOpt
    Set O.Dic = New Dictionary

Dim J%, K$
For J = 0 To L123_UB(A)
    With Eval_L123(A(J), Dic)
        If .Som Then
            O.Som = True
            K = L123_Key(A(J))
            O.Dic.Add K, .Str
        Else
            L123_Push O.L123AyLeft, A(J)
        End If
    End With
Next
Eval_L123Ay = O
End Function

Private Function Eval_L3Str(L3Str$, K$, Dic As Dictionary) As StrOpt
Dim L3 As L3: L3 = L3_Brk(L3Str)
If L3_HasSwitch(L3) Then
    If L3_SwitchIsOff(L3, Dic) Then Eval_L3Str = SomStr(""): Exit Function
End If
Dim A As KPD
    A = KPD(K, L3.Prm, Dic)
If IsSfx(K, ".Into") Then Stop
Dim O As StrOpt
    Select Case L3.Op
    Case eFixStr:   O = KPD_FixStr(L3.Prm)
    Case eFixDrp:   O = KPD_Drp(L3.Prm)
    Case eExp:      O = KPD_Exp(A)
    Case eExpIn:    O = KPD_ExpIn
    Case eFixFlag:  O = KPD_FixFlag
    Case eMac:      O = KPD_Mac(A)
    Case eRun:      O = KPD_Run(A)
    Case eExpComma: O = KPD_ExpComma(A)
    Case eSqlFix:   O = KPD_SqlFix
    Case eSqlPhrase: O = KPD_SqlPhrase(A)
    Case eFixLeftJn: O = KPD_FixLeftJn
    Case eFixAnd:    O = KPD_FixAnd
    Case eFixEq:     O = KPD_FixEq
    Case eFixOr:     O = KPD_FixOr
    Case eFixNe:     O = KPD_FixNe
    Case Else: Stop
    End Select
Eval_L3Str = O
End Function

Private Function FixOpAy() As eOp()

End Function

Private Function FixOpDic() As Dictionary

End Function

Private Function FixPrm_Val$(FixPrm$, FixOp)
End Function

Private Function FixPrm_Val_QQ$(FixPrm$, QQStr$)
End Function

Private Function Key_Ns$(K$)
Key_Ns = TakBefRev(K, ".")
End Function

Private Function KPD(K$, L3Prm$, Dic As Dictionary) As KPD
With KPD
    .K = K
    .P = L3Prm
    Set .D = Dic
End With
End Function

Private Function KPD_Drp(L3Prm$) As StrOpt
KPD_Drp = SomStr(L3Prm)
End Function

Private Function KPD_Exp(A As KPD) As StrOpt
Dim Sy$()
    With KPD_SyOpt(A)
        If Not .Som Then Exit Function
        Sy = .Sy
    End With
KPD_Exp = SomStr(JnVBar(Sy))
End Function

Private Function KPD_ExpComma(A As KPD) As StrOpt
Dim Ay$()
    Ay = SplitLvs(A.P)
Dim Sy$()
    With KPD_SyOpt(A)
        If Not .Som Then Exit Function
        Sy = .Sy
    End With
Dim O$()
    Dim J%, B$
    For J = 0 To UB(Ay)
        B = Sy(J) & " " & Ay(J)
        Push O, B
    Next
KPD_ExpComma = SomStr(Join(O, ",|    "))
End Function

Private Function KPD_ExpIn() As StrOpt

End Function

Private Function KPD_FixAnd() As StrOpt

End Function

Private Function KPD_FixEq() As StrOpt

End Function

Private Function KPD_FixFlag() As StrOpt

End Function

Private Function KPD_FixLeftJn() As StrOpt

End Function

Private Function KPD_FixNe() As StrOpt

End Function

Private Function KPD_FixOr() As StrOpt

End Function

Private Function KPD_FixStr(L3Prm$) As StrOpt
If FstChr(L3Prm) = "." Then
    KPD_FixStr = SomStr(RmvFstChr(L3Prm))
Else
    KPD_FixStr = SomStr(L3Prm)
End If
End Function

Private Function KPD_Mac(A As KPD) As StrOpt
Dim Sy$()
    With KPD_SyOpt(A)
        If Not .Som Then Exit Function
        Sy = .Sy
    End With
KPD_Mac = SomStr(JnVBar(Sy))
End Function

Private Function KPD_Run(A As KPD) As StrOpt
Dim Ay$()
    Ay = SplitLvs(A.P)
Dim FunNm$
    FunNm = AyShift(Ay)
Dim Av$()
    With PrmAy_SyOpt(A.K, Ay, A.D)
        If Not .Som Then Exit Function
        Av = .Sy
    End With
KPD_Run = SomStr(RunAv(FunNm, Av))
End Function

Private Function KPD_SqlFix() As StrOpt
Stop
End Function

Private Function KPD_SqlPhrase(A As KPD) As StrOpt
'Debug.Print "KPD_SqlPhrase-", A.P
Dim SqlKw$
SqlKw = TakAftRev(A.K, ".")
Dim O As StrOpt
Select Case SqlKw
Case "Sel": O = SqlPhrase_Sel(A)
Case "Fm": O = SqlPhrase_Fm(A.P)
Case "Into": O = SqlPhrase_Into(A.K)
Case "And"
Case "Gp"
Case "Upd": O = SqlPhrase_Upd(A.K)
Case "Set": O = SqlPhrase_Set(A)
Case "SelDis": O = SqlPhrase_SelDis(A)
Case "Wh": O = SqlPhrase_Wh(A)
Case "Jn": O = SqlPhrase_Jn(A)
Case Else
    Stop
End Select
KPD_SqlPhrase = O
End Function

Private Function KPD_SyOpt(A As KPD) As SyOpt
Dim Ay$(): Ay = SplitLvs(A.P)
KPD_SyOpt = PrmAy_SyOpt(A.K, Ay, A.D)
End Function

Private Function L123_Dr(A As L123) As Variant()
With A
    L123_Dr = Array(.LinI, .L1, .L2, .L3)
End With
End Function

Private Function L123_IsEq(A As L123, B As L123)
With A
    If .L1 <> B.L1 Then Exit Function
    If .L2 <> B.L2 Then Exit Function
End With
L123_IsEq = True
End Function

Private Function L123_IsPrmItm(A As L123) As Boolean
L123_IsPrmItm = A.L1 = "Prm"
End Function

Private Function L123_IsSwitchItm(A As L123) As Boolean
L123_IsSwitchItm = A.L1 = "?"
End Function

Private Function L123_Key$(A As L123)
L123_Key = A.L1 & "." & A.L2
End Function

Private Function L123_L1233(A As L123) As L1233
Dim O As L1233
With A
    O.L1 = .L1
    O.L2 = .L2
End With
With L3_L33(A.L3)
    O.Swtich = .Switch
    O.OpStr = .OpStr
    O.Prm = .Prm
End With
L123_L1233 = O
End Function

Private Sub L123_Push(Ay() As L123, M As L123)
Dim N%: N = L123_Sz(Ay)
ReDim Preserve Ay(N)
Ay(N) = M
End Sub

Private Sub L123_PushL3(O() As L123, I As L123)
Dim Idx%
    Idx = L123Ay_Idx(O, I)
If Idx = -1 Then
    L123_Push O, I
    Exit Sub
End If
Dim M As L123
    M = O(Idx)
    M.L3 = StrAppCrLf(M.L3, I.L3)
    O(Idx) = M
End Sub

Private Function L123_Sz%(Ay() As L123)
On Error Resume Next
L123_Sz = UBound(Ay) + 1
End Function

Private Function L123_UB%(Ay() As L123)
L123_UB = L123_Sz(Ay) - 1
End Function

Private Function L123_WrkDr(A As L123) As WrkDr
Dim O As WrkDr
With O
    O.LinI = A.LinI
    O.Ns = A.L1
    O.Nm = A.L2
    O.L3 = L3_Brk(A.L3)
End With
L123_WrkDr = O
End Function

Private Function L1233_Dr(A As L1233) As Variant()
With A
    L1233_Dr = Array(.LinI, .L1, .L2, .Swtich, .OpStr, .Prm)
End With
End Function

Private Function L1233Ay_Drs(A() As L1233) As Drs
Dim O As Drs
O.Fny = SplitSpc("LinI L1 L2 Switch OpStr Prm")
O.Dry = L1233Ay_Dry(A)
L1233Ay_Drs = O
End Function

Private Function L1233Ay_Dry(A() As L1233) As Variant()
Dim O(), J%
For J = 0 To UBound(A)
    Push O, L1233_Dr(A(J))
Next
L1233Ay_Dry = O
End Function

Private Sub L123Ay_Brw(A() As L123)
DrsWs L123Ay_Drs(A)
End Sub

Private Sub L123Ay_Brw_L1233(A() As L123)
DrsBrw L123Ay_Drs_L1233(A)
End Sub

Private Function L123Ay_Dic(A() As L123) As Dictionary
Dim O As Dictionary
    Set O = L123Ay_PrmDic(A)
    Set O = DicAdd(O, L123Ay_SwitchDic(A, O))
Dim B() As L123
    B = L123Ay_RmvPrmItm(A)
    B = L123Ay_RmvSwitchItm(B)
Dim J%
Do
    J = J + 1
    If J > 100 Then Stop
    Dim C As ExpOpt
    C = Eval_L123Ay(B, O)
    Set O = DicAdd(O, C.Dic)
    If Not C.Som Then Exit Do
    B = C.L123AyLeft
Loop
If Not L123Ay_IsEmpty(C.L123AyLeft) Then
    Dim Ay$()
        Ay = AyAdd(DicLy(O), L123Ay_Ly(C.L123AyLeft))
    Push Ay, "# of loops: " & J
    AyBrw Ay
    Stop
End If
Set L123Ay_Dic = O
End Function

Private Function L123Ay_DistOpSy(A() As L123) As String()
Dim J%, O() As eOp
For J = 0 To L123_UB(A)
    Push O, L3_Brk(A(J).L3).Op
Next
Dim OO$(), I, Op As eOp
For Each I In AyDist(O)
    Op = I
    Push OO, OpStr(Op)
Next
L123Ay_DistOpSy = OO
End Function

Private Function L123Ay_Drs(A() As L123) As Drs
Dim J%, Dry()
For J = 0 To L123_UB(A)
    Push Dry, L123_Dr(A(J))
Next
L123Ay_Drs.Dry = Dry
L123Ay_Drs.Fny = SplitSpc("LinI L1 L2 L3")
End Function

Private Function L123Ay_Drs_L1233(A() As L123) As Drs
Dim B() As L1233: B = L123Ay_L1233Ay(A)
L123Ay_Drs_L1233 = L1233Ay_Drs(B)
End Function

Private Function L123Ay_Idx%(A() As L123, I As L123)
Dim J%
For J = 0 To L123_UB(A)
    If L123_IsEq(A(J), I) Then L123Ay_Idx = J: Exit Function
Next
L123Ay_Idx = -1
End Function

Private Function L123Ay_IsEmpty(A() As L123) As Boolean
L123Ay_IsEmpty = L123_Sz(A) = 0
End Function

Private Function L123Ay_L1233Ay(A() As L123) As L1233()
Dim U%
    U = UBound(A)
Dim O() As L1233
    ReDim O(U)
Dim J%
For J = 0 To UBound(O)
    O(J) = L123_L1233(A(J))
Next
L123Ay_L1233Ay = O
End Function

Private Function L123Ay_L123Pass2Ay(A() As L123) As L123()
Dim U%
    U = UBound(A)
Dim O() As L123
    ReDim O(U)
    Dim J%
    Dim LasL1$, LasL2$, M As L123, I As L123
    Dim L1$, L2$, L3$, LinI%
    For J = 0 To U
        I = A(J)
        L1 = I.L1
        L2 = I.L2
        L3 = I.L3
        LinI = I.LinI
        If L3 <> "" Then
            If L1 = "" Then L1 = LasL1
            If L2 = "" Then L2 = LasL2
        ElseIf L2 <> "" Then
            If L1 = "" Then L1 = LasL1
        ElseIf L1 <> "" Then
        Else
            Stop
        End If
        M = L123(LinI, L1, L2, Trim(L3))
        O(J) = M
        LasL1 = M.L1
        LasL2 = M.L2
    Next
L123Ay_L123Pass2Ay = L123Ay_RmvL1(O)
End Function

Private Function L123Ay_L123Pass3Ay(A() As L123) As L123()
Dim J%
Dim O() As L123
For J = 0 To L123_UB(A)
    L123_PushL3 O, A(J)
Next
L123Ay_L123Pass3Ay = O
End Function

Private Function L123Ay_Ly(A() As L123) As String()
L123Ay_Ly = DrsLy(L123Ay_Drs(A))
End Function

Private Function L123Ay_PrmDic(A() As L123) As Dictionary
Dim O As New Dictionary
    Dim B() As L123: B = L123Ay_SelPrm(A)
    Dim J%
    Dim K$, V$
    For J = 0 To L123_UB(B)
        With B(J)
            V = L3_Brk(.L3).Prm
            K = .L1 & "." & .L2
            O.Add K, V
        End With
    Next
Set L123Ay_PrmDic = O
End Function

Private Function L123Ay_RmvL1(A() As L123) As L123()
Dim O() As L123, J%
For J = 0 To L123_UB(A)
    With A(J)
        If .L2 <> "" Or .L3 <> "" Then L123_Push O, A(J)
    End With
Next
L123Ay_RmvL1 = O
End Function

Private Function L123Ay_RmvPrmItm(A() As L123) As L123()
Dim J%, O() As L123
For J = 0 To L123_UB(A)
    If Not L123_IsPrmItm(A(J)) Then
        L123_Push O, A(J)
    End If
Next
L123Ay_RmvPrmItm = O
End Function

Private Function L123Ay_RmvSwitchItm(A() As L123) As L123()
Dim J%, O() As L123
For J = 0 To L123_UB(A)
    If Not L123_IsSwitchItm(A(J)) Then
        L123_Push O, A(J)
    End If
Next
L123Ay_RmvSwitchItm = O
End Function

Private Function L123Ay_SelL1(A() As L123, L1$) As L123()
Dim O() As L123, J%
For J = 0 To L123_UB(A)
    If A(J).L1 = L1 Then
        L123_Push O, A(J)
    End If
Next
L123Ay_SelL1 = O
End Function

Private Function L123Ay_SelPrm(A() As L123) As L123()
L123Ay_SelPrm = L123Ay_SelL1(A, "Prm")
End Function

Private Function L123Ay_SelSwitch(A() As L123) As L123()
L123Ay_SelSwitch = L123Ay_SelL1(A, "?")
End Function

Private Function L123Ay_SwitchDic(A() As L123, PrmDic As Dictionary) As Dictionary
Dim B() As L123
    B = L123Ay_SelSwitch(A)
Dim O As New Dictionary
    Dim Dic As Dictionary
    Set Dic = DicClone(PrmDic)
    Dim J%, K$, V As Boolean
    For J = 0 To L123_UB(B)
        K = "?" & B(J).L2
        V = L3Str_SwitchVal(B(J).L3, Dic)
        O.Add K, V
        Dic.Add K, V
    Next
Set L123Ay_SwitchDic = O
End Function

Private Function L123Ay_Wy(A() As L123) As WrkDr()
Dim U%
Dim O() As WrkDr
    U = UBound(A)
    ReDim O(U)
Dim J%, M As WrkDr
For J = 0 To U
    O(J) = L123_WrkDr(A(J))
Next
L123Ay_Wy = O
End Function

Private Function L3_Brk(L3$) As L3
Dim L$: L = Trim(L3)
If L3 = "" Then Exit Function
Dim Switch$
    If FstChr(L) = "?" Then
        Switch = FstTerm(L)
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

Private Function L3_HasSwitch(L3 As L3) As Boolean
L3_HasSwitch = L3.Switch <> ""
End Function

Private Function L3_Key$(A As L123)
L3_Key = A.L1 & "." & A.L2
End Function

Private Function L3_L33(L3) As L33
Dim O As L33
Dim A$: A = Trim(L3)
With O
    If FstChr(L3) = "?" Then
        .Switch = ParseTerm(A)
    End If
    .OpStr = ParseTerm(A)
    .Prm = A
End With
L3_L33 = O
End Function

Private Function L3_SwitchIsOff(L3 As L3, Dic As Dictionary) As Boolean
Dim V: V = Dic(L3.Switch)
If Not IsBool(V) Then Stop
L3_SwitchIsOff = Not V
End Function

Private Function L3Str_SwitchVal(L3Str$, Dic As Dictionary) As Boolean
Dim A As L3: A = L3_Brk(L3Str)
Select Case A.Op
Case eFixEq:  L3Str_SwitchVal = SwitchPrm_Val_Eq(A.Prm, Dic)
Case eFixNe:  L3Str_SwitchVal = SwitchPrm_Val_Ne(A.Prm, Dic)
Case eFixOr:  L3Str_SwitchVal = SwitchPrm_Val_Or(A.Prm, Dic)
Case eFixAnd: L3Str_SwitchVal = SwitchPrm_Val_And(A.Prm, Dic)
Case Else: Stop
End Select
End Function

Private Function LookupTerm_IsValid(K$, LookupTerm, Dic As Dictionary) As Boolean
LookupTerm_IsValid = True
If FstChr(LookupTerm) = "." Then Exit Function
If LookupTerm = "Into" Then Exit Function
If HasSubStr(LookupTerm, ".") Then
    If Dic.Exists(LookupTerm) Then Exit Function
End If
If Dic.Exists(K & "." & LookupTerm) Then Exit Function
LookupTerm_IsValid = False
End Function

Private Function LookupTerm_Val$(K$, LookupTerm, Dic As Dictionary)
If FstChr(LookupTerm) = "." Then
    LookupTerm_Val = LookupTerm
    Exit Function
End If
If LookupTerm = "Into" Then
    LookupTerm_Val = "  Into #" & TakAftRev(K, ".")
    Exit Function
End If
If HasSubStr(LookupTerm, ".") Then
    LookupTerm_Val = Dic(LookupTerm)
    Exit Function
End If
LookupTerm_Val = Dic(K & "." & LookupTerm)
End Function

Private Function Macro_ExpLis(Dic As Dictionary, Ns$, Nm$, TermLis) As SyOpt
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
Macro_ExpLis = SomSy(O)
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
Case "!": O = eRun
Case "$In": O = eExpIn
Case "#.": O = eSqlFix
Case "#": O = eSqlPhrase
Case "$": O = eMac
Case "$And": O = eMacAnd
Case "$Or": O = eExpOr
Case "$Wh": O = eMacWh
Case ".": O = eFixStr
Case ".And": O = eFixAnd
Case ".Comma": O = eFixComma
Case ".Drp": O = eFixDrp
Case ".Eq": O = eFixEq
Case ".Flag": O = eFixFlag
Case ".Fm": O = eFixFm
Case ".Gp": O = eFixGp
Case ".Jn": O = eFixJn
Case ".LeftJn": O = eFixLeftJn
Case ".Nbr": O = eNbr
Case ".NbrLis": O = eFixNbrLis
Case ".Ne": O = eFixNe
Case ".Or": O = eFixOr
Case ".Sel": O = eFixSel
Case ".SelDis": O = eFixSelDis
Case ".Set": O = eFixSet
Case ".Sql": O = eFixSql
Case ".StrLis": O = eFixStrLis
Case ".Upd": O = eFixUpd
Case ".Wh": O = eFixWh
Case "@": O = eExp
Case "@And": O = eExpAnd
Case "@Comma": O = eExpComma
Case "@Drp": O = eExpDrp
Case "@Gp": O = eExpGp
Case "@Jn": O = eExpJn
Case "@Jn": O = eExpJn
Case "@LeftJn": O = eExpLeftJn
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
For Each I In Array(eOp.eFixFm, eOp.eFixGp, eOp.eExpGp, eOp.eFixSelDis, eOp.eExpSelDis, eOp.eMac, eOp.eExp, eOp.eExpComma, _
    eOp.eFixLeftJn, eOp.eFixJn, eOp.eExpJn, eOp.eExpLeftJn, _
    eOp.eExpIn, eOp.eExpSel, eOp.eFixSel, eOp.eMacAnd, eOp.eMacOr, eOp.eFixStr)
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
    eOp.eExpIn, _
    eOp.eExpComma, _
    eOp.eExpDrp, _
    eOp.eExpGp, _
    eOp.eExpJn, _
    eOp.eExpOr, _
    eOp.eExpSel, _
    eOp.eExpSelDis, _
    eOp.eExpSet, _
    eOp.eExp
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

Private Function OpAy_FixXXX() As eOp()
Dim O() As eOp
Push O, eOp.eFixFm
Push O, eOp.eFixWh
Push O, eOp.eFixJn
Push O, eOp.eFixLeftJn
Push O, eOp.eFixComma
Push O, eOp.eFixDrp
Push O, eOp.eFixEq
Push O, eOp.eFixOr
Push O, eOp.eFixSel
Push O, eOp.eFixSelDis
OpAy_FixXXX = O

End Function

Private Function OpStr(A As eOp)
Dim O$
Select Case A
Case eExpIn: O = "$In"
Case eRun: O = "!"
Case eSqlFix: O = "#."
Case eSqlPhrase: O = "#"
Case eFixSql: O = ".Sql"
Case eFixEq:        O = ".Eq"
Case eExpAnd:    O = "@And"
Case eExpComma:  O = "@Comma"
Case eExpDrp:    O = "@Drp"
Case eExpGp:     O = "@Gp"
Case eExpJn:     O = "@Jn"
Case eExpLeftJn: O = "@LeftJn"
Case eExpOr:     O = "@Or"
Case eExpSel:    O = "@Sel"
Case eExpSelDis: O = "@SelDis"
Case eExpSet:    O = "@Set"
Case eExp:   O = "@"
Case eExpWh: O = "@Wh"
Case eFixAnd: O = ".And"
Case eFixComma:   O = ".Comma"
Case eFixDrp: O = ".Drp"
Case eFixJn: O = ".Jn"
Case eFixFm: O = ".Fm"
Case eFixGp: O = ".Gp"
Case eFixJn: O = ".Jn"
Case eFixOr: O = ".Or"
Case eFixSel: O = ".Sel"
Case eFixSelDis: O = ".SelDis"
Case eFixSet: O = ".Set"
Case eFixStr: O = "."
Case eFixUpd: O = ".Upd"
Case eFixWh: O = ".Wh"
Case eFixFlag: O = ".Flag"
Case eFixLeftJn:   O = ".LeftJn"
Case eMac:   O = "$"
Case eMacAnd: O = "$And"
Case eMacWh: O = "$Wh"
Case eMacOr:   O = "$Or"
Case eNbr: O = ".Nbr"
Case eFixNbrLis: O = ".NbrLis"
Case eFixNe: O = ".Ne"
Case eFixStrLis: O = ".StrLis"
Case Else: O = "?Unknown"
End Select
OpStr = O
End Function

Private Function PrmAy_SyOpt(K$, PrmAy$(), Dic As Dictionary) As SyOpt
'{Prm} in PrmAy is either with [.] or not.
'If FstChr is [.], just the {Prm}
'If with [.], just use {Prm} to lookup value in Dic
'If no [.], use {K}.{Prm} to lookup value in Dic
Dim LookupTerm
    For Each LookupTerm In PrmAy
        If Not LookupTerm_IsValid(K, LookupTerm, Dic) Then Exit Function
    Next
Dim Vy$()
    For Each LookupTerm In PrmAy
        Push Vy, LookupTerm_Val(K, LookupTerm, Dic)
    Next
PrmAy_SyOpt = SomSy(Vy)
End Function

Private Function PrmR_Val$(A As PrmR, PrmDic As Dictionary)
Dim ValAy(): ValAy = AyMap(A.PrmAy, "PrmRTerm_Val", PrmDic)
PrmR_Val = RunAv(A.FunNm, ValAy)
End Function

Private Function PrmRTerm_Val(PrmRTerm$, PrmDic As Dictionary) As String()
End Function

Private Function Soml123(LinI%, L1$, L2$, L3$) As L123Opt
Soml123.Som = True
Soml123.L123 = L123(LinI%, L1, L2, L3)
End Function

Private Function SomWrkDr(A As WrkDr) As WrkDrOpt
SomWrkDr.Som = True
SomWrkDr.WrkDr = A
End Function

Private Function Sql3_Rmv2Dash(Ly$()) As String()
Dim O$(), I
For Each I In Ly
    Push O, Brk1(I, "--", NoTrim:=True).S1
Next
Sql3_Rmv2Dash = O
End Function

Private Sub Sql3Ft_Rmv3Dash(Ft)
Dim Ly$(): Ly = FtLy(Ft)
Dim Ly1$(): Ly1 = Sql3Ly_No3DashLy(Ly)
If AyIsEq(Ly, Ly1) Then Exit Sub
AyWrt Ly1, Ft
End Sub

Private Function Sql3Ft_WrtEr(Ft) As Boolean
Sql3Ft_Rmv3Dash Ft
Dim Ly$(): Ly = FtLy(Ft)
Dim Ly1$(): Ly1 = Sql3Ly_ValidatedLy(Ly): If AyIsEmpty(Ly1) Then Exit Function
If AyIsEq(Ly, Ly1) Then Exit Function
AyWrt Ly1, Ft
Sql3Ft_WrtEr = True
End Function

Private Function Sql3Ly_AddEr(Sql3Ly$(), ErDry()) As String()
Dim I&, Dr
Dim W%, O$()
    O = Sql3Ly_No3DashLy(Sql3Ly)
    W = AyWdt(O)
For Each Dr In ErDry
    I = Dr(0)
    O(I) = AlignL(O(I), W) & " --- " & Dr(1)
Next
AyRmvEmptyEleAtEnd O
Push O, FmtQQ("--- [?] error(s)", Sz(ErDry))
Sql3Ly_AddEr = O
End Function

Private Function Sql3Ly_ErDry(Sql3Ly$()) As Variant()
Sql3Ly_ErDry = Wy_ErDry(Sql3Ly_Wy(Sql3Ly))
End Function

Private Function Sql3Ly_L123Ay(Sql3Ly$()) As L123()
Dim A() As L123: A = Sql3Ly_L123Pass1Ay(Sql3Ly)
Dim B() As L123: B = L123Ay_L123Pass2Ay(A)
Dim C() As L123: C = L123Ay_L123Pass3Ay(B)
Sql3Ly_L123Ay = C
End Function

Private Function Sql3Ly_L123Pass1Ay(Sql3Ly$()) As L123()
Dim O() As L123
    Dim L, LinI%, A As L123Opt
    For Each L In Sql3Ly_TrmLy(Sql3Ly)
        With TrmLin_L123Opt(L, LinI)
            If .Som Then
                L123_Push O, .L123
            End If
        End With
        LinI = LinI + 1
    Next
Sql3Ly_L123Pass1Ay = O
End Function

Private Function Sql3Ly_LinLvlDrs(Sql3Ly$()) As Drs
Dim Dry(), Dr
Dim L
For Each L In Sql3Ly
    Push Dry, Array(TrmLin_Lvl(L), L)
Next
Dim O As Drs
    O.Fny = SplitSpc("Lvl Lin")
    O.Dry = Dry
Sql3Ly_LinLvlDrs = O
End Function

Private Function Sql3Ly_No3DashLy(Sql3Ly$()) As String()
Sql3Ly_No3DashLy = AyMapIntoSy(Sql3Ly, "Lin_No3DashLin")
End Function

Private Function Sql3Ly_ValidatedLy(Sql3Ly$()) As String()
Dim ErDry(): ErDry = Sql3Ly_ErDry(Sql3Ly)
If AyIsEmpty(ErDry) Then Exit Function
Sql3Ly_ValidatedLy = Sql3Ly_AddEr(Sql3Ly, ErDry)
End Function

Private Function Sql3Ly_Wy(Sql3Ly$()) As WrkDr()
Dim A() As L123: A = Sql3Ly_L123Pass1Ay(Sql3Ly)
Dim B() As L123: B = L123Ay_L123Pass2Ay(A)
Dim C() As L123: C = L123Ay_L123Pass3Ay(B)
Sql3Ly_Wy = L123Ay_Wy(B)
End Function

Private Function SqlPhrase_Fm(L3Prm$) As StrOpt
SqlPhrase_Fm = SomStr("|  From " & L3Prm)
End Function

Private Function SqlPhrase_Into(K$) As StrOpt
Dim T$
    With BrkRev(K$, ".")
        If .S2 <> "Into" Then Stop
        T = TakAftRev(.S1, ".")
    End With
SqlPhrase_Into = SomStr("|  Into " & T)
End Function

Private Function SqlPhrase_Jn(A As KPD) As StrOpt
Dim Ay$()
    Ay = SplitLvs(A.P)
Dim Sy$()
    With KPD_SyOpt(A)
        If Not .Som Then Exit Function
        Sy = .Sy
    End With
Dim O$()
    O = AyAddPfx("|    ", AyRmvEmpty(Sy))
SqlPhrase_Jn = SomStr(Join(O))
End Function

Private Function SqlPhrase_Sel(A As KPD) As StrOpt
Dim Ay$()
    Ay = SplitLvs(A.P)
Dim Sy$()
    With KPD_SyOpt(A)
        If Not .Som Then Exit Function
        Sy = .Sy
    End With
Dim O$()
    Dim J%, B$
    For J = 0 To UB(Ay)
        If FstChr(Ay(J)) = "." Then
            B = RmvFstChr(Ay(J))
        ElseIf FstChr(Ay(J)) = "?" Then
            If Sy(J) = "" Then GoTo Nxt
            Ay(J) = RmvFstChr(Ay(J))
            B = Sy(J) & " " & Ay(J)
        Else
            B = Sy(J) & " " & Ay(J)
        End If
        Push O, B
Nxt:
    Next
SqlPhrase_Sel = SomStr("Select|    " & Join(O, ",|    "))
End Function

Private Function SqlPhrase_SelDis(A As KPD) As StrOpt
Dim Ay$()
    Ay = SplitLvs(A.P)
Dim Sy$()
    With KPD_SyOpt(A)
        If Not .Som Then Exit Function
        Sy = .Sy
    End With
Dim O$()
    Dim J%, B$
    For J = 0 To UB(Ay)
        B = Sy(J) & " " & Ay(J)
        Push O, B
    Next
SqlPhrase_SelDis = SomStr("Select Distinct|    " & Join(O, ",|    "))
Stop
End Function

Private Function SqlPhrase_Set(A As KPD) As StrOpt
Dim Ay$()
    Ay = SplitLvs(A.P)
Dim Sy$()
    With KPD_SyOpt(A)
        If Not .Som Then Exit Function
        Sy = .Sy
    End With
Dim O$()
    Dim J%, B$
    For J = 0 To UB(Ay)
        B = Ay(J) & " = " & Sy(J)
        Push O, B
    Next
SqlPhrase_Set = SomStr("  Set|    " & Join(O, ",|    "))
End Function

Private Function SqlPhrase_Upd(K$) As StrOpt
Dim Ns$: Ns = Key_Ns(K)
Dim A$: A = TakAftRev(Ns, ".")
Dim T$: T = TakBefRev(A, "#")
Dim O$
    O = "Update #" & T
SqlPhrase_Upd = SomStr(O)
End Function

Private Function SqlPhrase_Wh(A As KPD) As StrOpt
Dim B$
    
Dim O$
    O = "|  Where " & B
SqlPhrase_Wh = SomStr(O)
End Function

Private Function Sts_StsExp(A As Sts) As Sts
Dim O As Sts
'    Dim Wy() As WrkDr
'    Sts_SplitOp A, eOp.eExp, O, Wy
'    Dim J%, M As Wr
'    For J = 0 To UB(A.Wrs.Dry)
'        M = Wr(A.Wrs.Dry(J))
'        Dim K$, V$
'        K = M.Ns & "." & M.Nm
'        V = Wy_StsExp__Val(M.Prm)
'        O.Dic.Add K, V
'    Next
Sts_StsExp = O
End Function

Private Function Sts_StsSwitch(A As Sts1) As Sts1
Dim O As Sts1
    Dim Wrs As Drs
    Sts1_SplitNs A, "?", O, Wrs
    Dim J%, M As Wr, K$
    For J = 0 To UB(Wrs.Dry)
        M = Wrs_Wr(Wrs, J)
        K = "?" & M.Nm
        O.Dic.Add K, SwitchWr_Val(M, O.Dic) '<===
    Next
Sts_StsSwitch = O
End Function

Private Sub Sts1_SplitNs(A As Sts1, Ns$, O As Sts1, ByRef OWrs As Drs)
Dim IdxAy&()
    IdxAy = Wrs_IdxAy_Ns(A.Wrs, Ns)
OWrs = DrsSelRow(A.Wrs, IdxAy)
Set O.Dic = DicClone(A.Dic)
O.Wrs = DrsExlRow(A.Wrs, IdxAy)
End Sub

Private Sub Sts1_SplitOp(A As Sts1, Op As eOp, O As Sts1, OWrs As Drs)
Dim IdxAy&()
    IdxAy = Wrs_IdxAy_Op(A.Wrs, eFixStr)
OWrs = DrsSelRow(A.Wrs, IdxAy)
Set O.Dic = DicClone(A.Dic)
O.Wrs = DrsExlRow(A.Wrs, IdxAy)
End Sub

Private Function Sts1_StsFixStr(A As Sts1) As Sts1
Dim O As Sts1
    Dim Wrs As Drs
    Sts1_SplitOp A, eFixStr, O, Wrs
    Dim J%
    Dim M As Wr
    For J = 0 To UB(Wrs.Dry)
        M = Wrs_Wr(Wrs, J)
        Dim K$: K = M.Ns & "." & M.Nm
        Dim V$
            If SwitchVal(A.Dic, M.Switch) Then
                V = M.Prm
            Else
                V = ""
            End If
        O.Dic.Add K, V
    Next
Sts1_StsFixStr = O
End Function

Private Function Sts1_StsPrm(A As Sts1) As Sts1
Dim O As Sts1
    Dim Wrs As Drs
    Sts1_SplitNs A, "Prm", O, Wrs
    Dim J%
    Dim K$, V$
    For J = 0 To UB(Wrs.Dry)
        With Wrs.Dry(J)
            V = .Prm
            K = .Ns & "." & .Nm
            O.Dic.Add K, V           '<====
        End With
    Next
Sts1_StsPrm = O
End Function

Private Function Sts1Pair_Ds(Bef As Sts1, Aft As Sts1, DsNm$) As Ds
Dim O As Ds
Dim Cur() As WrkDr
'    Cur = Wy_Minus(Bef.Wy, Aft.Wy)
Dim CurDic As Dictionary
    Set CurDic = DicMinus(Bef.Dic, Aft.Dic)
'DsAddDt O, Dt("Bef", Wy_Drs(Bef.Wy, Aft.Dic))
'DsAddDt O, Dt("Aft", Wy_Drs(Aft.Wy, Aft.Dic))
'DsAddDt O, Dt("Cur", Wy_Drs(Cur, Aft.Dic))
'DsAddDt O, DicDt(Bef.Dic, "BefDic")
'DsAddDt O, DicDt(Aft.Dic, "AftDic")
'DsAddDt O, DicDt(CurDic, "CurDic")
Sts1Pair_Ds = O
End Function

Private Sub StsPair_Assert(Bef As Sts, Aft As Sts)
Dim Diff%
    Dim B%, A%
    B = Wy_UB(Bef.Wy)
    A = Wy_UB(Aft.Wy)
    Diff = A - B
Debug.Assert Diff <> Aft.Dic.Count - Bef.Dic.Count
End Sub

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

Private Sub SwitchNm_Assert(SwitchNm$)
If FstChr(SwitchNm) <> "?" Then Stop
End Sub

Private Function SwitchPrm_V1V2(SwitchPrm$, Dic As Dictionary) As S1S2
SwitchTerm_Assert SwitchPrm
Dim T1$, T2$
    With Brk(SwitchPrm, " ")
        T1 = .S1
        T2 = .S2
    End With
    
Dim V1$
    V1 = DicVal(Dic, T1)
    If V1 = "{?}" Then
        Stop
        Exit Function
    End If
Dim V2$
    If T2 = "*Blank" Then V2 = "" Else V2 = T2
SwitchPrm_V1V2 = S1S2(V1, V2)
End Function

Private Sub SwitchTerm_Assert(SwitchTerm)
Dim A$: A = FstChr(SwitchTerm)
If A = "?" Then Exit Sub
If IsPfx(SwitchTerm, "Prm.") Then Exit Sub
Stop
End Sub

Private Function SwitchTerm_Val(SwitchTerm$, Dic As Dictionary) As BoolOpt
SwitchTerm_Assert SwitchTerm
If FstChr(SwitchTerm) = "?" Then
    SwitchTerm_Val = DicBoolOpt(Dic, SwitchTerm)
    Exit Function
End If
SwitchTerm_Val = DicBoolOpt(Dic, SwitchTerm)
End Function

Private Function SwitchVal(SwitchDic As Dictionary, SwitchNm$) As Boolean
If SwitchNm = "" Then SwitchVal = True: Exit Function
SwitchNm_Assert SwitchNm
DicAssertKey SwitchDic, SwitchNm
SwitchVal = SwitchDic(SwitchNm)
End Function

Private Function SwitchValStr$(SwitchDic As Dictionary, SwitchNm$)
If SwitchNm = "" Then Exit Function
If IsNothing(SwitchDic) Then
    SwitchValStr = "{?}"
    Exit Function
End If
With DicBoolOpt(SwitchDic, SwitchNm)
    If .Som Then
        SwitchValStr = .Bool
    Else
        SwitchValStr = "{?}"
    End If
End With
End Function

Private Function SwitchWr_Val(A As Wr, Dic As Dictionary) As Boolean
Select Case A.Op
Case eFixEq:  SwitchWr_Val = SwitchPrm_Val_Eq(A.Prm, Dic)
Case eFixNe:  SwitchWr_Val = SwitchPrm_Val_Ne(A.Prm, Dic)
Case eFixOr:  SwitchWr_Val = SwitchPrm_Val_Or(A.Prm, Dic)
Case eFixAnd: SwitchWr_Val = SwitchPrm_Val_And(A.Prm, Dic)
Case Else: Stop
End Select
End Function

Private Function SwitchWrkDr_Val(A As WrkDr, Dic As Dictionary) As Boolean
Select Case A.L3.Op
Case eFixEq:  SwitchWrkDr_Val = SwitchPrm_Val_Eq(A.L3.Prm, Dic)
Case eFixNe:  SwitchWrkDr_Val = SwitchPrm_Val_Ne(A.L3.Prm, Dic)
Case eFixOr:  SwitchWrkDr_Val = SwitchPrm_Val_Or(A.L3.Prm, Dic)
Case eFixAnd: SwitchWrkDr_Val = SwitchPrm_Val_And(A.L3.Prm, Dic)
Case Else: Stop
End Select
End Function

Private Function TrmLin_IsL1(L) As Boolean
Dim C$
C = FstChr(L)
TrmLin_IsL1 = IsLetter(C) Or C = "?"
End Function

Private Function TrmLin_IsL2(L) As Boolean
If Left(L, 4) = Space(4) Then
    Dim C$: C = Mid(L, 5, 1)
    TrmLin_IsL2 = IsLetter(C) Or C = "?"
End If
End Function

Private Function TrmLin_IsL3(L) As Boolean
If Left(L, 8) = Space(8) Then
    TrmLin_IsL3 = Mid(L, 9, 1) <> " "
End If
End Function

Private Function TrmLin_L123Opt(TrmLin, LinI%) As L123Opt
Dim O As L123Opt
Dim Lvl%: Lvl = TrmLin_Lvl(TrmLin)
If Not IsIn(Lvl, 1, 2, 3) Then Exit Function
Dim L1$, L2$, L3$, A$
Select Case Lvl
    Case 1: L1 = TrmLin                ' Assume L1 can only have one term
    Case 2: A = TrmLin: L2 = ParseTerm(A): L3 = A  ' Assume L2 can have (L2) or (L2 & L3)
    Case 3: L3 = TrmLin
End Select
TrmLin_L123Opt = Soml123(LinI%, L1, L2, L3)
End Function

Private Function TrmLin_Lvl(L) As Byte
If TrmLin_IsL1(L) Then TrmLin_Lvl = 1: Exit Function
If TrmLin_IsL2(L) Then TrmLin_Lvl = 2: Exit Function
If TrmLin_IsL3(L) Then TrmLin_Lvl = 3: Exit Function
End Function

Private Function Wr(Dr) As Wr
With Wr
    AyAsg Dr, .LinI, .Ns, .Nm, .Switch, .Op, .Prm
End With
End Function

Private Function WrkDr_Key$(A As WrkDr)
WrkDr_Key = A.Ns & "." & A.Nm
End Function

Private Function Wrs_IdxAy_Ns(Wrs As Drs, Ns$) As Long()
Dim INs%
    INs = AyIdx(Wrs.Fny, "Ns")
Dim O&()
    Dim J%
    For J = 0 To UB(Wrs.Dry)
        If Wrs.Dry(J)(INs) = Ns Then
            Push O, J
        End If
    Next
Wrs_IdxAy_Ns = O
End Function

Private Function Wrs_IdxAy_Op(Wrs As Drs, Op As eOp) As Long()
Dim IOp%
    IOp = AyIdx(Wrs.Fny, "Op")
Dim O&()
    Dim J%
    For J = 0 To UB(Wrs.Dry)
        If Wrs.Dry(J)(IOp) = Op Then
            Push O, J
        End If
    Next
Wrs_IdxAy_Op = O
End Function

Private Function Wrs_Wr(Wrs As Drs, J%) As Wr
Wrs_Wr = Wr(Wrs.Dry(J))
End Function

Private Function Wy_BefAftCurDs(Bef() As WrkDr, Aft() As WrkDr, Optional Dic As Dictionary) As Ds
Dim O As Ds
Dim Done() As WrkDr
    Done = Wy_Minus(Bef, Aft)
DsAddDt O, Dt("Bef", Wy_Drs(Bef, Dic))
DsAddDt O, Dt("Aft", Wy_Drs(Aft, Dic))
DsAddDt O, Dt("Cur", Wy_Drs(Done, Dic))
Wy_BefAftCurDs = O
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

Private Function Wy_Dr(A As WrkDr, Dic As Dictionary) As Variant()
Dim SwitchV$
    SwitchV = SwitchValStr(Dic, A.L3.Switch)
With A
    Wy_Dr = Array(.LinI, .Ns, .Nm, .L3.Switch, SwitchV, OpStr(.L3.Op), .L3.Prm, .L3.L3)
End With
End Function

Private Function Wy_Drs(Wy() As WrkDr, Optional Dic As Dictionary) As Drs
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

Private Function Wy_ErDry(Wy() As WrkDr) As Variant()
Dim O()
PushAy O, Wy_ErDry_InvalidOp(Wy)
PushAy O, Wy_ErDry_NotAlwSwitch(Wy)
PushAy O, Wy_ErDry_SwitchNotExist(Wy)
PushAy O, Wy_ErDry_UpdMstHavNamWithPondSign(Wy)
PushAy O, Wy_ErDry_NoPrm(Wy)
PushAy O, Wy_ErDry_NoSql(Wy)
Wy_ErDry = O
End Function

Private Function Wy_ErDry_InvalidOp(Wy() As WrkDr) As Variant()
Dim J%
Dim O()
For J = 0 To UBound(Wy)
    With Wy(J).L3
        If .Op = eOp.eUnknown Then
            Stop
            Push O, Array(Wy(J).LinI, FmtQQ("Invalid Op[?]", .OpStr))
        End If
    End With
Next
Wy_ErDry_InvalidOp = O
End Function

Private Function Wy_ErDry_NoPrm(Wy() As WrkDr) As Variant()
Dim J%
For J = 0 To Wy_UB(Wy)
    If Wy(J).Ns = "Prm" Then Exit Function
Next
Wy_ErDry_NoPrm = Array(Array(0, "Warning: No Prml namespace"))
End Function

Private Function Wy_ErDry_NoSql(Wy() As WrkDr) As Variant()
Dim J%
For J = 0 To Wy_UB(Wy)
    If Wy(J).Ns = "Sql" Then Exit Function
Next
Wy_ErDry_NoSql = Array(Array(0, "Warning: No Sql namespace"))
End Function

Private Function Wy_ErDry_NotAlwSwitch(Wy() As WrkDr) As Variant()
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
Wy_ErDry_NotAlwSwitch = O
End Function

Private Function Wy_ErDry_SwitchNotExist(Wy() As WrkDr) As Variant()
Dim J%, O()
For J = 0 To UBound(Wy)
    With Wy(J).L3
        If .Switch = "" Then GoTo Nxt
        If Switch_Exist(Wy, .Switch) Then GoTo Nxt
        Push O, Array(J, "Switch not exist")
    End With
Nxt:
Next
Wy_ErDry_SwitchNotExist = O
End Function

Private Function Wy_ErDry_UpdMstHavNamWithPondSign(Wy() As WrkDr) As Variant()

End Function

Private Function Wy_Fny() As String()
Wy_Fny = SplitSpc("LinI Ns Nm Switch SwitchVal Op Prm L3")
End Function

Private Function Wy_Has(A() As WrkDr, B As WrkDr) As Boolean
Dim J%
For J = 0 To Wy_UB(A)
    If A(J).LinI = B.LinI Then Wy_Has = True: Exit Function
Next
End Function

Private Function Wy_IdxAy_Ky(Wy() As WrkDr, Ky$()) As Integer()
Dim J%, O%()
For J = 0 To Wy_UB(Wy)
    Dim K$
    With Wy(J)
        K = .Ns & "." & .Nm
    End With
    If AyHas(Ky, K) Then Push O, J
Next
Wy_IdxAy_Ky = O
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

Private Function Wy_Ky(Wy() As WrkDr) As String()
Dim J%, O$()
For J = 0 To Wy_UB(Wy)
    Dim K$
    With Wy(J)
        Push O, .Ns & "." & .Nm
    End With
Next
Wy_Ky = O
End Function

Private Function Wy_L3Dic_FixStr(Wy() As WrkDr) As Dictionary
Dim O As New Dictionary
Dim K$, V$
Dim J%, M As WrkDr
For J = 0 To Wy_UB(Wy)
    M = Wy(J)
    If M.L3.Op = eFixStr Then
        K = M.Ns & "." & M.Nm
        V = M.L3.L3
        O.Add K, V
    End If
Next
Set Wy_L3Dic_FixStr = O
End Function

Private Function Wy_L3Dic_Switch(Wy() As WrkDr) As Dictionary
Dim O As New Dictionary
Dim K$, V$
Dim J%, M As WrkDr
For J = 0 To Wy_UB(Wy)
    M = Wy(J)
    If M.Ns = "?" Then
        K = "?" & M.Nm
        V = M.L3.L3
        O.Add K, V
    End If
Next
Set Wy_L3Dic_Switch = O
End Function

Private Function Wy_Minus(A() As WrkDr, B() As WrkDr) As WrkDr()
Dim O() As WrkDr
Dim J%
For J = 0 To Wy_UB(A)
    If Not Wy_Has(B, A(J)) Then Wy_Push O, A(J)
Next
Wy_Minus = O
End Function

Private Function Wy_MulNmIdxAy(Wy() As WrkDr, Ky$()) As Integer()

End Function

Private Function Wy_MulNmKy(Wy() As WrkDr) As String()
Wy_MulNmKy = AyDupAy(Wy_Ky(Wy))
End Function

Private Function Wy_MulNmVal$(Key, Wy() As WrkDr)

End Function

Private Function Wy_MulNmWy(Wy() As WrkDr) As WrkDr()
Dim Ky$()
    Ky = Wy_MulNmKy(Wy)
Dim IdxAy%()
    IdxAy = Wy_IdxAy_Ky(Wy, Ky)
Wy_MulNmWy = Wy_Sel(Wy, IdxAy)
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

Private Function Wy_Sel(Wy() As WrkDr, IdxAy%()) As WrkDr()
If AyIsEmpty(IdxAy) Then Exit Function
Dim U%: U = UB(IdxAy)
Dim O() As WrkDr
ReDim O(U)
Dim I, J%
For Each I In IdxAy
    O(J) = Wy(I)
    J = J + 1
Next
Wy_Sel = O
End Function

Private Function Wy_StsExp(Wy() As WrkDr) As StsExp
Dim IdxAy%(): IdxAy = Wy_IdxAy_Op(Wy, eOp.eExp)
Dim ODic As New Dictionary
    Dim J%, M As WrkDr
    For J = 0 To UB(IdxAy)
        M = Wy(IdxAy(J))
        Dim K$, V$
        K = M.Ns & "." & M.Nm
        V = Wy_StsExp__Val(M.L3.Prm)
        ODic.Add K, V
    Next
Set Wy_StsExp.ExpDic = ODic
Wy_StsExp.RestWy = Wy_RmvItms(Wy, IdxAy)
End Function

Private Function Wy_StsExp__Val$(Prm$)
Wy_StsExp__Val$ = Prm
End Function

Private Function Wy_StsFixDrp(Wy() As WrkDr) As StsFixDrp
Dim IdxAy%(): IdxAy = Wy_IdxAy_Op(Wy, eOp.eFixDrp)
Dim ODic As New Dictionary
    Dim J%, M As WrkDr
    For J = 0 To UB(IdxAy)
        M = Wy(IdxAy(J))
        Dim K$, V$
        K = M.Ns & "." & M.Nm
        V = Wy_StsFixDrp__Sqs(M.L3.Prm)
        ODic.Add K, V
    Next
Set Wy_StsFixDrp.DrpDic = ODic
Wy_StsFixDrp.RestWy = Wy_RmvItms(Wy, IdxAy)
End Function

Private Function Wy_StsFixDrp__Sqs$(TblNmLvs$)
Wy_StsFixDrp__Sqs$ = TblNmLvs
End Function

Private Function Wy_StsFixFm(Wy() As WrkDr) As StsFixFm
Dim IdxAy%(): IdxAy = Wy_IdxAy_Op(Wy, eFixFm)
Dim W() As WrkDr: W = Wy_Sel(Wy, IdxAy)
Dim ODic As New Dictionary
    Dim J%
    Dim M As WrkDr
    For J = 0 To Wy_UB(W)
        M = W(J)
        If Trim(M.L3.Prm) = "" Then Stop
        Dim K$: K = M.Ns & "." & M.Nm & ".Fm"
        Dim V$: V = "|  From " & M.L3.Prm
        ODic.Add K, V
    Next
Wy_StsFixFm.RestWy = Wy_RmvItms(Wy, IdxAy)
Set Wy_StsFixFm.FmDic = ODic
End Function

Private Function Wy_StsFixJn(Wy() As WrkDr) As StsFixJn
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
Set Wy_StsFixJn.JnDic = ODic
Wy_StsFixJn.RestWy = Wy_RmvItms(Wy, IdxAy)
End Function

Private Function Wy_StsFixLeftJn(Wy() As WrkDr) As StsFixLeftJn
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
Set Wy_StsFixLeftJn.LeftJnDic = ODic
Wy_StsFixLeftJn.RestWy = Wy_RmvItms(Wy, IdxAy)
End Function

Private Function Wy_StsFixSelDis(Wy() As WrkDr) As StsFixSelDis
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
Set Wy_StsFixSelDis.SelDisDic = ODic
Wy_StsFixSelDis.RestWy = Wy_RmvItms(Wy, IdxAy)
End Function

Private Function Wy_StsFixStr(Wy() As WrkDr, SwitchDic As Dictionary) As StsFixStr
Dim IdxAy%():                   IdxAy = Wy_IdxAy_Op(Wy, eFixStr)
Dim ODic As New Dictionary
    Dim J%
    Dim M As WrkDr
    For J = 0 To UB(IdxAy)
        M = Wy(IdxAy(J))
        Dim K$: K = M.Ns & "." & M.Nm
        Dim V$
            If SwitchVal(SwitchDic, M.L3.Switch) Then
                V = M.L3.Prm
            Else
                V = ""
            End If
        ODic.Add K, V
    Next
Set Wy_StsFixStr.StrDic = ODic
Wy_StsFixStr.RestWy = Wy_RmvItms(Wy, IdxAy)
End Function

Private Function Wy_StsFixUpd(Wy() As WrkDr) As StsFixUpd
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
Set Wy_StsFixUpd.UpdDic = ODic
Wy_StsFixUpd.RestWy = Wy_RmvItms(Wy, IdxAy)
End Function

Private Function Wy_StsFixWh(Wy() As WrkDr) As StsFixWh
Dim IdxAy%(): IdxAy = Wy_IdxAy_Op(Wy, eFixWh)
Dim ODic As New Dictionary
    Dim J%, M As WrkDr
    For J = 0 To UB(IdxAy)
        M = Wy(IdxAy(J))
        Dim K$: K = M.Ns & "." & M.Nm & ".Where"
        Dim V$: V = "|  Where " & M.L3.Prm
        ODic.Add K, V
    Next
Set Wy_StsFixWh.WhDic = ODic
Wy_StsFixWh.RestWy = Wy_RmvItms(Wy, IdxAy)
End Function

Private Function Wy_StsFixXXX(Wy() As WrkDr, Op As eOp) As StsFixXXX
Dim IdxAy%(): IdxAy = Wy_IdxAy_Op(Wy, Op)
Dim ODic As New Dictionary
    Dim J%
    Dim M As WrkDr
    Dim OpDic As Dictionary
        Set OpDic = FixOpDic
    For J = 0 To UB(IdxAy)
        M = Wy(IdxAy(J))
        If Trim(M.L3.Prm) = "" Then Stop
        Dim K$: K = M.Ns & "." & M.Nm & ".Fm"
        Dim V$: V = FixPrm_Val(M.L3.Prm, OpDic(M.L3.Op))
        ODic.Add K, V
    Next
Wy_StsFixXXX.RestWy = Wy_RmvItms(Wy, IdxAy)
Set Wy_StsFixXXX.XXXDic = ODic
End Function

Private Function Wy_StsFixXXXAll(Wy() As WrkDr) As StsFixXXX
Dim I, Op As eOp
Dim IWy() As WrkDr
Dim A As StsFixXXX
For Each I In FixOpAy
    Op = I
    A = Wy_StsFixXXX(IWy, Op)
Next
End Function

Private Function Wy_StsMulNm(Wy() As WrkDr) As StsMulNm
Dim Ky$()
    Ky = Wy_MulNmKy(Wy)
Dim ODic As New Dictionary
    Dim K
    For Each K In Ky
        ODic.Add K, Wy_MulNmVal(K, Wy)
    Next
Dim IdxAy%()
    IdxAy = Wy_MulNmIdxAy(Wy, Ky)
Set Wy_StsMulNm.MulNmDic = ODic
Wy_StsMulNm.RestWy = Wy_RmvItms(Wy, IdxAy)
End Function

Private Function Wy_StsSwitch(A As StsPrm) As StsSwitch
Dim Wy() As WrkDr
Dim PrmDic As Dictionary
    Wy = A.RestWy
    Set PrmDic = A.PrmDic
Dim IdxAy%()
    IdxAy = Wy_IdxAy_Ns(Wy, "?")
Dim ODic As Dictionary
    Set ODic = A.PrmDic
    Dim J%, M As WrkDr, K$
    For J = 0 To UB(IdxAy)
        M = Wy(IdxAy(J))
        K = "?" & M.Nm
        ODic.Add K, SwitchWrkDr_Val(M, ODic)
    Next
Set Wy_StsSwitch.SwitchDic = ODic
Wy_StsSwitch.RestWy = Wy_RmvItms(Wy, IdxAy)
End Function

Private Function Wy_Sz&(Wy() As WrkDr)
On Error Resume Next
Wy_Sz = UBound(Wy) + 1
End Function

Private Function Wy_UB&(Wy() As WrkDr)
Wy_UB = Wy_Sz(Wy) - 1
End Function

Private Function ZZSql3Ft$()
ZZSql3Ft = TstResPth & "SalRpt.Sql3"
End Function

Private Sub ZZSql3Ft_Fix()
Dim O$(): O = ZZSql3Ly
Dim J%
For J = 0 To UB(O)
    O(J) = Replace(O(J), Chr(160), " ")
Next
AyWrt O, ZZSql3Ft
End Sub

Private Function ZZSql3Ly() As String()
ZZSql3Ly = FtLy(ZZSql3Ft)
End Function

Private Function ZZSql3Ly_L123Ay() As L123()
ZZSql3Ly_L123Ay = Sql3Ly_L123Ay(ZZSql3Ly)
End Function

Private Function ZZSql3Ly_Wy() As WrkDr()
ZZSql3Ly_Wy = Sql3Ly_Wy(ZZSql3Ly)
End Function

Private Function ZZSts() As Sts
ZZSts.Wy = ZZWy
Set ZZSts.Dic = New Dictionary
End Function

Private Function ZZWy() As WrkDr()
If Sql3Ft_WrtEr(ZZSql3Ft) Then ZZSql3Ft_Edt: Stop
ZZWy = Sql3Ly_Wy(ZZSql3Ly)
End Function

Private Sub KPD_Run__Tst()
Dim Dic As New Dictionary
Dim L3Prm$
Dim Ns$
    Ns = ""
    L3Prm = "CrdExpr Prm.CrdTyLis"
    Dic.Add "Prm.CrdTyLis", "1 2 3"
Dim A As KPD
    A = KPD(Ns, L3Prm, Dic)
With KPD_Run(A)
    Debug.Assert .Som
    Debug.Print .Str
End With
End Sub

Private Sub L123Ay_Brw_L1233__Tst()
Dim A() As L123: A = Sql3Ly_L123Pass1Ay(ZZSql3Ly)
Dim B() As L123: B = L123Ay_L123Pass2Ay(A)
Dim C() As L123: C = L123Ay_L123Pass3Ay(B)
L123Ay_Brw_L1233 C
End Sub

Private Sub L123Ay_DistOpSy__Tst()
Dim A() As L123: A = Sql3Ly_L123Pass1Ay(ZZSql3Ly)
Dim B() As L123: B = L123Ay_L123Pass2Ay(A)
Dim C() As L123: C = L123Ay_L123Pass3Ay(B)
AyBrw L123Ay_DistOpSy(C)
End Sub

Private Sub L123Ay_L123Pass2Ay__Tst()
Dim A() As L123: A = Sql3Ly_L123Pass1Ay(ZZSql3Ly)
Dim B() As L123: B = L123Ay_L123Pass2Ay(A)
L123Ay_Brw B
End Sub

Private Sub L123Ay_L123Pass3Ay__Tst()
Dim A() As L123: A = Sql3Ly_L123Pass1Ay(ZZSql3Ly)
Dim B() As L123: B = L123Ay_L123Pass2Ay(A)
Dim C() As L123: C = L123Ay_L123Pass3Ay(B)
L123Ay_Brw C
End Sub

Private Sub Op_Sy__Tst()
AyDmp Op_Sy
End Sub

Sub Sql3Ft_Dic__Tst()
If Sql3Ft_WrtEr(ZZSql3Ft) Then
    FtBrw ZZSql3Ft
    Stop
End If
DicBrw Sql3Ft_Dic(ZZSql3Ft)
End Sub

Private Sub Sql3Ft_WrtWy_ErDry__Tst()
If Sql3Ft_WrtEr(ZZSql3Ft) Then ZZSql3Ft_Edt
End Sub

Private Sub Sql3Ly_L123Pass1Ay__Tst()
L123Ay_Brw Sql3Ly_L123Pass1Ay(ZZSql3Ly)
End Sub

Private Sub Sql3Ly_LinLvlDrs__Tst()
DrsBrw Sql3Ly_LinLvlDrs(ZZSql3Ly)
End Sub

Private Sub Sql3Ly_TrmLy__Tst()
AyBrw Sql3Ly_TrmLy(ZZSql3Ly)
End Sub

Private Sub Sql3Ly_ValidatedLy__Tst()
Dim Ly$(): Ly = Sql3Ly_ValidatedLy(ZZSql3Ly)
If AyIsEmpty(Ly) Then Exit Sub
AyWrt Ly, ZZSql3Ft
ZZSql3Ft_Edt
End Sub

Private Sub Sql3Ly_Wy__Tst()
Sql3Ft_Rmv3Dash ZZSql3Ft
Dim Act() As WrkDr: Act = Sql3Ly_Wy(ZZSql3Ly)
Wy_Brw Act
End Sub

Private Sub SqlPhrase_Upd__Tst()
With SqlPhrase_Upd("Sql.Tx.Tx#Upd.Upd")
    Debug.Assert .Str = "Update #Tx"
    Debug.Assert .Som = True
End With
End Sub

Private Sub Sts_StsFixStr__Tst()
Dim A As Sts: A = ZZSts
Dim B As Sts: 'B = Sts_StsPrm(A)
Dim X As Sts: 'X = Sts_StsSwitch(B)
Dim Y As Sts: 'Y = Sts_StsFixStr(X)
StsPair_Assert X, Y
'DsWbVis Sts1Pair_Ds(X, Y, "Sts_StsFixStr")
End Sub

Private Function Sts_StsMulNm__Tst()
Const A$ = "Prm Switch FixStr MulNm"
Dim FunNy$(): FunNy = AyAddPfx(SplitLvs(A), "Sts_Sts")
Dim FunNy1$(): FunNy1 = FunNy: AyRmvLasEle FunNy
Dim X As Sts: 'X = Sts_StsFixStr(Sts_StsSwitch(Sts_StsPrm(ZZSts)))
Dim Y As Sts: 'Y = Sts_StsPrm(X)
Dim Dif%
    Dim Bef%: Bef = Wy_UB(X.Wy)
    Dim Aft%: Aft = Wy_UB(Y.Wy)
    Dif = Bef - Aft
Debug.Assert Y.Dic.Count - X.Dic.Count = Dif
Wy_Brw X.Wy
End Function

Private Function Sts_StsPrm__Tst()
Dim A As Sts: A = ZZSts
Dim B As Sts: 'B = Sts_StsPrm(A)  '<== Run
Dim Diff%
    Dim Bef%, Aft%
    Bef = Wy_UB(A.Wy)
    Aft = Wy_UB(B.Wy)
    Diff = Bef - Aft
Debug.Assert B.Dic.Count - A.Dic.Count = Diff
Dim J%
For J = 0 To Wy_UB(B.Wy)
    With B.Wy(J)
        Debug.Assert .Ns <> "Prm"
    End With
Next
Dim O As Ds
    O = Wy_BefAftCurDs(A.Wy, B.Wy)
    O.DsNm = "Wy_StsPrm Result"
    DsAddDt O, DicDt(B.Dic, "Dic")
DsWbVis O
End Function

Private Function Sts_StsSwitch__Tst()
Dim A As Sts: 'A = Sts_StsPrm(ZZSts)
Dim B As Sts: 'B = Sts_StsSwitch(A) '<== Run
Dim Diff%
    Dim Bef%, Aft%
    Bef = Wy_UB(A.Wy)
    Aft = Wy_UB(B.Wy)
    Diff = Bef - Aft
Debug.Assert B.Dic.Count - A.Dic.Count = Diff
Dim J%
For J = 0 To Wy_UB(B.Wy)
    With B.Wy(J)
        Debug.Assert .Ns <> "Prm"
    End With
Next
Dim O As Ds
    O = Wy_BefAftCurDs(A.Wy, B.Wy)
    O.DsNm = "Wy_StsPrm Result"
    DsAddDt O, DicDt(A.Dic, "PrmDic")
DsWbVis O
End Function

Private Sub Wy_MulNmWy__Tst()
Wy_Brw Wy_MulNmWy(ZZWy)
End Sub

Private Sub Wy_StsExp__Tst()
Dim Wy() As WrkDr: Wy = ZZWy
Dim A As StsPrm: A = Wy_StsPrm(Wy)
Dim B As StsSwitch: B = Wy_StsSwitch(A)
Dim C As StsFixStr: '' C = Wy_StsFxStr(B)
Dim D As StsFixFm: D = Wy_StsFixFm(C.RestWy)
Dim F As StsFixWh: F = Wy_StsFixWh(D.RestWy)
Dim G As StsFixUpd: G = Wy_StsFixUpd(F.RestWy)
Dim H As StsFixLeftJn: H = Wy_StsFixLeftJn(G.RestWy)
Dim I As StsFixJn: I = Wy_StsFixJn(H.RestWy)
Dim J As StsFixSelDis: J = Wy_StsFixSelDis(I.RestWy)
Dim K As StsFixDrp: K = Wy_StsFixDrp(J.RestWy)
Dim L As StsExp: L = Wy_StsExp(K.RestWy)
Dim Dif%
    Dim Bef%: Bef = Wy_UB(K.RestWy)
    Dim Aft%: Aft = Wy_UB(L.RestWy)
    Dif = Bef - Aft
Debug.Assert L.ExpDic.Count = Dif
Dim O As Ds
    O = Wy_BefAftCurDs(K.RestWy, L.RestWy, B.SwitchDic)
Dim DD As Dictionary
    'Const DicPfx$ = "Prm Switch FixStr Fm Into Wh Upd LeftJn Jn SelDis Drp Exp"
    'Set DD = DicMge(DicPfx, A.PrmDic, B.SwitchDic, C.FixStrDic, D.FmDic, E.IntoDic, F.WhDic, G.UpdDic, H.LeftJnDic, I.JnDic, J.SelDisDic, K.DrpDic, L.ExpDic)
    Set DD = DicAdd(A.PrmDic, B.SwitchDic, C.StrDic, D.FmDic, F.WhDic, G.UpdDic, H.LeftJnDic, I.JnDic, J.SelDisDic, K.DrpDic, L.ExpDic)
    If DicHasBlankKey(B.SwitchDic) Then Stop
    If DicHasBlankKey(C.StrDic) Then Stop
DsAddDt O, DicDt(DD, "Mged Dic")
DsAddDt O, DicDt(L.ExpDic, "Exp Dic (@)")
O.DsNm = "Wy_StsExp Result"
DsWbVis O
End Sub

Private Sub Wy_StsFixDrp__Tst()
Dim Wy() As WrkDr: Wy = ZZWy
Dim A As StsPrm: A = Wy_StsPrm(Wy)
Dim B As StsSwitch: B = Wy_StsSwitch(A)
Dim C As StsFixStr: ' C = Wy_StsFxStr(B)
Dim D As StsFixFm: D = Wy_StsFixFm(C.RestWy)
Dim F As StsFixWh: F = Wy_StsFixWh(D.RestWy)
Dim G As StsFixUpd: G = Wy_StsFixUpd(F.RestWy)
Dim H As StsFixLeftJn: H = Wy_StsFixLeftJn(G.RestWy)
Dim I As StsFixJn: I = Wy_StsFixJn(H.RestWy)
Dim J As StsFixSelDis: J = Wy_StsFixSelDis(I.RestWy)
Dim K As StsFixDrp: K = Wy_StsFixDrp(J.RestWy)
Dim Dif%
    Dim Bef%: Bef = Wy_UB(J.RestWy)
    Dim Aft%: Aft = Wy_UB(K.RestWy)
    Dif = Bef - Aft
Debug.Assert K.DrpDic.Count = Dif
Dim O As Ds
    O = Wy_BefAftCurDs(J.RestWy, K.RestWy, B.SwitchDic)
    DsAddDt O, DicDt(K.DrpDic, "Drp Dic")
O.DsNm = "Wy_StsFixDrp Result"
DsWbVis O
End Sub

Private Sub Wy_StsFixFm__Tst()
Dim Wy() As WrkDr: Wy = ZZWy
Dim A As StsPrm: A = Wy_StsPrm(Wy)
Dim B As StsSwitch: B = Wy_StsSwitch(A)
Dim C As StsFixStr: ' C = Wy_StsFxStr(B)
Dim D As StsFixFm: D = Wy_StsFixFm(C.RestWy)
Dim Dif%
    Dim Aft%: Aft = Wy_UB(D.RestWy)
    Dim Bef%: Bef = Wy_UB(C.RestWy)
    Dif = Bef - Aft
Debug.Assert Dif = D.FmDic.Count
Dim O As Ds
    O = Wy_BefAftCurDs(C.RestWy, D.RestWy, B.SwitchDic)
    DsAddDt O, DicDt(D.FmDic, "Fm Dic")
DsWbVis O
End Sub

Private Sub Wy_StsFixJn__Tst()
Dim Wy() As WrkDr: Wy = ZZWy
Dim A As StsPrm: A = Wy_StsPrm(Wy)
Dim B As StsSwitch: B = Wy_StsSwitch(A)
Dim C As StsFixStr: ' C = Wy_StsFxStr(B)
Dim D As StsFixFm: D = Wy_StsFixFm(C.RestWy)
Dim F As StsFixWh: F = Wy_StsFixWh(D.RestWy)
Dim G As StsFixUpd: G = Wy_StsFixUpd(F.RestWy)
Dim H As StsFixLeftJn: H = Wy_StsFixLeftJn(G.RestWy)
Dim I As StsFixJn: I = Wy_StsFixJn(H.RestWy)
DicBrw I.JnDic
Dim Dif%
    Dim Bef%: Bef = Wy_UB(H.RestWy)
    Dim Aft%: Aft = Wy_UB(I.RestWy)
    Dif = Bef - Aft
Debug.Assert I.JnDic.Count = Dif
Wy_Brw I.RestWy, B.SwitchDic
End Sub

Private Sub Wy_StsFixLeftJn__Tst()
Dim Wy() As WrkDr: Wy = ZZWy
Dim A As StsPrm: A = Wy_StsPrm(Wy)
Dim B As StsSwitch: B = Wy_StsSwitch(A)
Dim C As StsFixStr: ' C = Wy_StsFxStr(B)
Dim D As StsFixFm: D = Wy_StsFixFm(C.RestWy)
Dim F As StsFixWh: F = Wy_StsFixWh(D.RestWy)
Dim G As StsFixUpd: G = Wy_StsFixUpd(F.RestWy)
Dim H As StsFixLeftJn: H = Wy_StsFixLeftJn(G.RestWy)
DicBrw H.LeftJnDic
Dim Dif%
    Dim Bef%: Bef = Wy_UB(G.RestWy)
    Dim Aft%: Aft = Wy_UB(H.RestWy)
    Dif = Bef - Aft
Debug.Assert H.LeftJnDic.Count = Dif
Wy_Brw H.RestWy, B.SwitchDic
End Sub

Private Sub Wy_StsFixSelDis__Tst()
Dim Wy() As WrkDr: Wy = ZZWy
Dim A As StsPrm: A = Wy_StsPrm(Wy)
Dim B As StsSwitch: B = Wy_StsSwitch(A)
Dim C As StsFixStr: ' C = Wy_StsFxStr(B)
Dim D As StsFixFm: D = Wy_StsFixFm(C.RestWy)
Dim F As StsFixWh: F = Wy_StsFixWh(D.RestWy)
Dim G As StsFixUpd: G = Wy_StsFixUpd(F.RestWy)
Dim H As StsFixLeftJn: H = Wy_StsFixLeftJn(G.RestWy)
Dim I As StsFixJn: I = Wy_StsFixJn(H.RestWy)
Dim K As StsFixSelDis: K = Wy_StsFixSelDis(I.RestWy)
DicBrw K.SelDisDic
Dim Dif%
    Dim Bef%: Bef = Wy_UB(I.RestWy)
    Dim Aft%: Aft = Wy_UB(K.RestWy)
    Dif = Bef - Aft
Debug.Assert K.SelDisDic.Count = Dif
Wy_Brw K.RestWy, B.SwitchDic
End Sub

Private Sub Wy_StsFixStr__Tst()
Dim Wy() As WrkDr: Wy = ZZWy
Dim A As StsPrm: A = Wy_StsPrm(Wy)
Dim B As StsSwitch: B = Wy_StsSwitch(A)
Dim C As StsFixStr: ' C = Wy_StsFxStr(B)
Dim Dif%
    Dim Bef%: Bef = Wy_UB(B.RestWy)
    Dim Aft%: Aft = Wy_UB(C.RestWy)
    Dif = Bef - Aft
Debug.Assert C.StrDic.Count = Dif
Dim O As Ds
    O = Wy_BefAftCurDs(B.RestWy, C.RestWy, B.SwitchDic)
    O.DsNm = "Wy_StsFixStr Result"
    DsAddDt O, DicDt(C.StrDic, "FixStrDic")
    DsAddDt O, DicDt(B.SwitchDic, "SwitchDic")
    DsAddDt O, DicDt(Wy_L3Dic_FixStr(B.RestWy), "FixStrL3Dic")
DsWbVis O
End Sub

Private Sub Wy_StsFixUpd__Tst()
Dim Wy() As WrkDr: Wy = ZZWy
Dim A As StsPrm: A = Wy_StsPrm(Wy)
Dim B As StsSwitch: B = Wy_StsSwitch(A)
Dim C As StsFixStr: ' C = Wy_StsFxStr(B)
Dim D As StsFixFm: D = Wy_StsFixFm(C.RestWy)
Dim F As StsFixWh: F = Wy_StsFixWh(D.RestWy)
Dim G As StsFixUpd: G = Wy_StsFixUpd(F.RestWy)
DicBrw G.UpdDic
Dim Dif%
    Dim Bef%: Bef = Wy_UB(F.RestWy)
    Dim Aft%: Aft = Wy_UB(G.RestWy)
    Dif = Bef - Aft
Debug.Assert G.UpdDic.Count = Dif
Wy_Brw G.RestWy, B.SwitchDic
End Sub

Private Sub Wy_StsFixWh__Tst()
Dim Wy() As WrkDr: Wy = ZZWy
Dim A As StsPrm: A = Wy_StsPrm(Wy)
Dim B As StsSwitch: B = Wy_StsSwitch(A)
Dim C As StsFixStr: ' C = Wy_StsFxStr(B)
Dim D As StsFixFm: D = Wy_StsFixFm(C.RestWy)
Dim F As StsFixWh: F = Wy_StsFixWh(D.RestWy)
DicBrw F.WhDic
Dim Dif%
    Dim Bef%: Bef = Wy_UB(D.RestWy)
    Dim Aft%: Aft = Wy_UB(F.RestWy)
    Dif = Bef - Aft
Debug.Assert F.WhDic.Count = Dif
Wy_Brw F.RestWy, B.SwitchDic
End Sub

Private Sub Wy_StsPrm__Tst()
Dim Wy() As WrkDr: Wy = ZZWy
Dim A As StsPrm: A = Wy_StsPrm(Wy) '<== Run
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
Dim O As Ds
    O = Wy_BefAftCurDs(Wy, A.RestWy)
    O.DsNm = "Wy_StsPrm Result"
    DsAddDt O, DicDt(A.PrmDic, "PrmDic")
DsWbVis O
End Sub

Sub Wy_StsSwitch__Tst()
Dim Wy() As WrkDr: Wy = ZZWy
Dim A As StsPrm: A = Wy_StsPrm(Wy)
Dim B As StsSwitch: B = Wy_StsSwitch(A)
Dim Diff%
    Dim Aft%, Bef%
    Bef = Wy_UB(A.RestWy)
    Aft = Wy_UB(B.RestWy)
    Diff = Bef - Aft
Debug.Assert B.SwitchDic.Count = Diff
Dim O As Ds
    O = Wy_BefAftCurDs(A.RestWy, B.RestWy, B.SwitchDic)
    O.DsNm = "Wy_StsSwitch Result"
    Dim DD As Dictionary
        Dim SwitchL3Dic As Dictionary
            Set SwitchL3Dic = Wy_L3Dic_Switch(Wy)
        Set DD = DicMge("SwtichL3Val SwitchVal Prm", SwitchL3Dic, B.SwitchDic, A.PrmDic)
    DsAddDt O, DicDt(DD, "Dic for checking")
DsWbVis O
End Sub

Private Sub ZZSql3Ly__Tst()
AyBrw ZZSql3Ly
End Sub

Sub Tst()
Wy_StsExp__Tst
Wy_StsFixDrp__Tst
Wy_StsFixFm__Tst
Wy_StsFixJn__Tst
Wy_StsFixLeftJn__Tst
Wy_StsFixSelDis__Tst
Wy_StsFixStr__Tst
Wy_StsFixUpd__Tst
Wy_StsFixWh__Tst
Wy_StsPrm__Tst
Wy_StsSwitch__Tst
Op_Sy__Tst
Sql3Ly_ValidatedLy__Tst
Sql3Ft_WrtWy_ErDry__Tst
Sql3Ly_Wy__Tst
Sql3Ft_Dic__Tst
ZZSql3Ly__Tst
End Sub

