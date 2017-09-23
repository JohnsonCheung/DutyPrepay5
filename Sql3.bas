Attribute VB_Name = "Sql3"
Option Explicit
Option Compare Database
Public Enum eOp
    eExpIn  ' [$In]
    eRun    ' [!]
    eSqlPhrase ' [#]
    eSqlStmt     ' [*]
    eSqlSel    ' [#Sel]
    eFixEq      ' [.EQ]
    eFixNe      ' [.NE]
    eExpComma   ' [@Comma]
    eExp  ' [@]   means <Prm> is term list for sql-statment
    eExpWh   ' [@Wh] means <Prm> is term list for "Sql-Where"
    eFixAnd  ' [.And] means <Prm> is fixed str for "Sql-And"
    eFixOr  ' [.Or] means <Prm> is fixed str for "Sql-Or"
    eFixComma ' [.Comma]
    eFixDrp
    eFixStr    ' [.]   means <Prm> is fixed str
    eMac       ' [$] means <Prm> is a macro string ( a template string with {..} to be expand.  Inside {..} is a <Ns>.<Nm>.
    eUnknown '
End Enum
Private Type ErDrOpt
    Som As Boolean
    ErDr() As Variant
End Type
Private Type L12SOPS
    LinI As Integer
    L1 As String
    L2 As String
    Op As eOp
    Switch As String
    Prm As String
    S As Variant  ' Must be Boolean for Switch or String for other
    Don As Boolean
End Type
Private Type L12SOPSAyOpt
    L12SOPSAy() As L12SOPS
    Som As Boolean
End Type
Private Type KPD
    K As String   ' Ns.Nm
    P As String   ' L3Prm
    D As Dictionary
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

Sub AA(Optional Sql3FtVarNm = "")
Evl__Tst Sql3FtVarNm
End Sub

Sub AAA()
Evl_ToDic__Tst
End Sub

Sub Edt(Optional Sql3FtVarNm = "")
ZZSql3Ft_Edt Sql3FtVarNm
End Sub

Private Function Evl(Sql3Ft$) As Dictionary
Dim Ly$(): Ly = FtLy(Sql3Ft)
Dim L123Ay() As L123
L123Ay = Evl_1____FmSql3Ly_ToL123Ay(Ly)
'L123Ay = L123Ay_SelNsPfxLvs(L123Ay, "Sql.X.T.Tx Expr Prm ?")
Set Evl = Evl_2____FmL123Ay_ToDic(L123Ay)
End Function

Private Function Evl_1____FmSql3Ly_ToL123Ay(Sql3Ly$()) As L123()
Dim A() As L123: A = Evl_1___Pass1(Sql3Ly)
Dim B() As L123: B = Evl_1___Pass2(A)
Evl_1____FmSql3Ly_ToL123Ay = B
End Function

Private Function Evl_1___Pass1(Sql3Ly$()) As L123()
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
Evl_1___Pass1 = O
End Function

Private Function Evl_1___Pass2(A() As L123) As L123()
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
Evl_1___Pass2 = Evl_1___Pass2_RmvL1(O)
End Function

Private Function Evl_1___Pass2_RmvL1(A() As L123) As L123()
Dim O() As L123, J%
For J = 0 To L123_UB(A)
    With A(J)
        If .L2 <> "" Or .L3 <> "" Then L123_Push O, A(J)
    End With
Next
Evl_1___Pass2_RmvL1 = O
End Function

Private Function Evl_2____FmL123Ay_ToDic(A() As L123) As Dictionary
Dim B1() As L12SOPS
Dim B2() As L12SOPS
Dim B3() As L12SOPS
    B1 = Evl_21___FmL123Ay_ToL12SOPSAy(A)
    B2 = Evl_22___EvlPrm(B1)
    B3 = Evl_23___EvlSwitch(B2)
Dim C1() As L12SOPS
    C1 = B3
Dim J%, C2() As L12SOPS
For J = 0 To 99
    C2 = Evl_24___EvlOneCycle(C1)
    
    Dim Bef_UB%: Bef_UB = L12SOPS_UB(C1)
    Dim Aft_UB%: Aft_UB = L12SOPS_UB(C2)
    If Aft_UB <> Bef_UB Then Stop
    
    Dim Bef%: Bef = L12SOPSAy_DonCnt(C1)
    Dim Aft%: Aft = L12SOPSAy_DonCnt(C2)
    If Bef = Aft Then GoTo X
    C1 = C2
Next
Never
X:
Dim O As Dictionary
    Set O = Evl_ToDic(C2)

Dim IsAllDon As Boolean: IsAllDon = L12SOPSAy_IsAllDon(C2)
Dim NotDonCnt%:          NotDonCnt = L12SOPSAy_NotDonCnt(C2)
If Not IsAllDon Then
    Dim Dt1 As Dt: Dt1 = L12SOPSAy_Dt(L12SOPSAy_Srt(C2))
                   Dt1 = DtDrpCol(Dt1, "L1 L2")
                   Dt1 = DtReOrd(Dt1, "Done")
    Dim Dt2 As Dt: Dt2 = DicDt(O, "Rslt Dic")
    Dim Dt3 As Dt: Dt3 = DicDt(DicByDry(Array( _
        Array("Loop Count", J), _
        Array("IsAllDon", IsAllDon), _
        Array("NotDonCnt", NotDonCnt))))
    Dim Ds As Ds
        DsAddDt Ds, Dt1
        DsAddDt Ds, Dt2
        DsAddDt Ds, Dt3
    DtBrw Dt1
    DtBrw Dt2
    DtBrw Dt3
    Stop
End If
Set Evl_2____FmL123Ay_ToDic = O
End Function

Private Function Evl_21___FmL123Ay_ToL12SOPSAy(A() As L123) As L12SOPS()
Dim O() As L12SOPS
Dim J%
For J = 0 To L123_UB(A)
    L12SOPS_Push O, L123_L12SOPS(A(J))
Next
Evl_21___FmL123Ay_ToL12SOPSAy = O
End Function

Private Function Evl_22___EvlPrm(A() As L12SOPS) As L12SOPS()
Dim O() As L12SOPS
    O = A
Dim J%
For J = 0 To L12SOPS_UB(A)
    If A(J).L1 = "Prm" Then
        O(J).S = A(J).Prm
        O(J).Don = True
    End If
Next
Evl_22___EvlPrm = O
End Function

Private Function Evl_23___EvlSwitch(A() As L12SOPS) As L12SOPS()
Dim O() As L12SOPS
    O = A
Dim X%
For X = 0 To 99
    Dim J%
    With Evl_231__EvlSwitchOneCycle(O)
        If .Som Then
            O = .L12SOPSAy
            GoTo Nxt
        End If
        If Not L12SOPSAy_IsAllSwitchSet(O) Then Stop
        Evl_23___EvlSwitch = O
        Exit Function
    End With
Nxt:
Next
Debug.Print "Never reach here"
Stop
End Function

Private Function Evl_23___EvlSwitchVal(Op As eOp, Prm$, Dic As Dictionary) As Boolean
Dim O As Boolean
Select Case Op
Case eFixEq:  O = Evl_231__EvlSwitchVal_ForEq(Prm, Dic)
Case eFixNe:  O = Evl_232__EvlSwitchVal_ForNe(Prm, Dic)
Case eFixOr:  O = Evl_233__EvlSwitchVal_ForOr(Prm, Dic)
Case eFixAnd: O = Evl_234__EvlSwitchVal_ForAnd(Prm, Dic)
Case Else: Stop
End Select
Evl_23___EvlSwitchVal = O
End Function

Private Function Evl_231__EvlSwitchOneCycle(A() As L12SOPS) As L12SOPSAyOpt
Dim J%, O() As L12SOPS, Som As Boolean
O = A
For J = 0 To L12SOPS_UB(A)
    With Evl_2311_EvlSwitchItm(O, J)
        If .Som Then
            O(J).S = .Bool      'Only case for .S to Be boolean
            O(J).Don = True
            Som = True
        End If
    End With
Next
If Som Then
    Evl_231__EvlSwitchOneCycle = SomL12SOPSAy(O)
End If
End Function

Private Function Evl_231__EvlSwitchVal_ForEq(SwitchPrm$, Dic As Dictionary) As Boolean
With SwitchPrm_V1V2(SwitchPrm$, Dic)
    Evl_231__EvlSwitchVal_ForEq = .S1 = .S2
End With
End Function

Private Function Evl_2311_EvlSwitchItm(A() As L12SOPS, Idx%) As BoolOpt
Dim M As L12SOPS
M = A(Idx)
If M.Don Then Exit Function
If M.L1 <> "?" Then Exit Function
If M.Switch <> "" Then Stop ' SwitchItm cannot have Switch
Dim L3Str$
Dim Dic As Dictionary
    Set Dic = Evl_ToDic(A)
    L3Str = OpStr(M.Op) & " " & M.Prm
Evl_2311_EvlSwitchItm = SomBool(L3Str_SwitchVal(L3Str, Dic))
End Function

Private Function Evl_232__EvlSwitchVal_ForNe(SwitchPrm$, Dic As Dictionary) As Boolean
With SwitchPrm_V1V2(SwitchPrm$, Dic)
    Evl_232__EvlSwitchVal_ForNe = .S1 <> .S2
End With
End Function

Private Function Evl_233__EvlSwitchVal_ForOr(SwitchPrm$, Dic As Dictionary) As Boolean
Evl_233__EvlSwitchVal_ForOr = BoolAy_Or(SwitchPrm_BoolAy(SwitchPrm$, Dic))
End Function

Private Function Evl_234__EvlSwitchVal_ForAnd(SwitchPrm$, Dic As Dictionary) As Boolean
Evl_234__EvlSwitchVal_ForAnd = BoolAy_Or(SwitchPrm_BoolAy(SwitchPrm$, Dic))
End Function

Private Function Evl_24___EvlOneCycle(A() As L12SOPS) As L12SOPS()
Dim O() As L12SOPS
    O = A
Dim J%
For J = 0 To L12SOPS_UB(A)
    With Evl_241__EvlOneIdx(A, J)
        If .Som Then
            O(J).S = .Str
            O(J).Don = True
        End If
    End With
Next
Evl_24___EvlOneCycle = O
End Function

Private Function Evl_241__EvlOneIdx(A() As L12SOPS, Idx%) As StrOpt
Dim M As L12SOPS: M = A(Idx)
If M.Don Then Exit Function
Dim K$: K = L12SOPS_Key(M)
Dim O() As L12SOPS
    O = A
Dim Dic As Dictionary
    Set Dic = Evl_ToDic(O)
Evl_241__EvlOneIdx = Evl_2412_EvlOneItm(K, M.Switch, M.Op, M.Prm, Dic)
End Function

Private Function Evl_2412_EvlOneItm(K$, Switch$, Op As eOp, Prm$, Dic As Dictionary) As StrOpt
If Evl_24120_HasSwitch_And_IsOff(Switch, Dic) Then
    Evl_2412_EvlOneItm = SomStr("")
    Exit Function
End If

Dim A As KPD
    A = KPD(K, Prm, Dic)
Dim O As StrOpt
    Select Case Op
    Case eSqlStmt:   O = Evl_24121_SqlStmt(A)
    Case eFixStr:    O = Evl_24122_FixStr(Prm)
    Case eFixDrp:    O = Evl_24123_Drp(Prm)
    Case eExp:       O = Evl_24124_Exp(A)
    Case eExpIn:     O = Evl_24125_ExpIn(A)
    Case eMac:       O = Evl_24126_Mac(A)
    Case eRun:       O = Evl_24127_Run(A)
    Case eExpComma:  O = Evl_24128_ExpComma(A)
    Case eSqlPhrase: O = Evl_24129_SqlPhrase(A)
    Case eSqlSel:    O = Evl_2412A_SqlSelFldLis(A)
    Case Else: Stop
    End Select
Evl_2412_EvlOneItm = O
End Function

Private Function Evl_24120_HasSwitch_And_IsOff(L3Switch$, Dic As Dictionary) As Boolean
If L3Switch = "" Then Exit Function
If Not Dic.Exists(L3Switch) Then Stop
Dim V: V = Dic(L3Switch)
If Not IsBool(V) Then Stop
Evl_24120_HasSwitch_And_IsOff = Not V
End Function

Private Function Evl_24121_SqlStmt(A As KPD) As StrOpt
Dim Sy$()
    'A.P is a list of sql phrase.  Some of them does not find in A.D, but they can evaluated directly
    'Those terms are: Into Upd
    Dim Ay$(): Ay = SplitLvs(A.P)
    Dim I
    For Each I In Ay
        Select Case I
        Case "Into":    Push Sy, "|  Into #" & RmvPfx(TakAftRev(A.K, "."), "?")
        Case "Upd":     Push Sy, "Update #" & TakAftRev(A.K, "#")
        Case Else
            With Evl_PrmAy_1__EvlOneTerm(A.K, I, A.D)
                If Not .Som Then Exit Function
                Push Sy, .Str
            End With
        End Select
    Next
Evl_24121_SqlStmt = SomStr(Join(Sy) & "|")
End Function

Private Function Evl_24122_FixStr(L3Prm$) As StrOpt
Evl_24122_FixStr = SomStr(RplFstChrToIf(L3Prm, ".", " "))
End Function

Private Function Evl_24123_Drp(L3Prm$) As StrOpt
Evl_24123_Drp = SomStr("Drop " & L3Prm & "|")
End Function

Private Function Evl_24124_Exp(A As KPD) As StrOpt
Dim Sy$()
    With Evl_PrmLvs(A)
        If Not .Som Then Exit Function
        Sy = .Sy
    End With
Evl_24124_Exp = SomStr(JnVBar(Sy))
End Function

Private Function Evl_24125_ExpIn(A As KPD) As StrOpt
Static N%
N = N + 1
Dim Ay$(): Ay = SplitLvs(A.P)
If Sz(Ay) <> 2 Then Stop
If N >= 20 Then Stop
Dim S$, InLis$
    With Evl_PrmAy(A.K, Ay, A.D)
        If Not .Som Then Exit Function
        S = .Sy(0)
        InLis = .Sy(1)
    End With
Evl_24125_ExpIn = SomStr(FmtQQ("? In (?)", S, InLis))
End Function

Private Function Evl_24126_Mac(A As KPD) As StrOpt
Dim Sy$()
    With Evl_PrmLvs(A)
        If Not .Som Then Exit Function
        Sy = .Sy
    End With
Evl_24126_Mac = SomStr(JnVBar(Sy))
End Function

Private Function Evl_24127_Run(A As KPD) As StrOpt
Dim Ay$()
    Ay = SplitLvs(A.P)
Dim MthNm$
    MthNm = AyShift(Ay)
Dim Av$()
    With Evl_PrmAy(A.K, Ay, A.D)
        If Not .Som Then Exit Function
        Av = .Sy
    End With
Evl_24127_Run = SomStr(RunAv(MthNm, Av))
End Function

Private Function Evl_24128_ExpComma(A As KPD) As StrOpt
Dim Ay$()
    Ay = SplitLvs(A.P)
Dim Sy$()
    With Evl_PrmLvs(A)
        If Not .Som Then Exit Function
        Sy = .Sy
    End With
Dim O$()
    Dim J%, B$
    For J = 0 To UB(Ay)
        B = Sy(J) & " " & Ay(J)
        Push O, B
    Next
Evl_24128_ExpComma = SomStr(Join(O, ",|    "))
End Function

Private Function Evl_24129_SqlPhrase(A As KPD) As StrOpt
'Debug.Print "Evl_24129_SqlPhrase-", A.P
Dim SqlKw$
SqlKw = TakAftRev(A.K, ".")
Dim O As StrOpt
Select Case SqlKw
Case "Sel":    O = Evl_241291__SqlSel(A)
Case "Into":   Stop: O = Evl_241292__SqlInto(A.K) ' No need to define [Into]
Case "Fm":     O = Evl_241293__SqlFm(A.P)
Case "And":    O = Evl_241294__SqlAnd(A)          ' No need to define [Into]
Case "Gp":     O = Evl_241295__SqlGp(A)
Case "Upd":    Stop: O = Evl_241296__SqlUpd(A.K)  ' No need to define [Upd]
Case "Set":    O = Evl_241297__SqlSet(A)
Case "SelDis": O = Evl_241298__SqlSelDis(A)
Case "Wh":     O = Evl_241299__SqlWh(A)
Case "Jn":     O = Evl_24129A__SqlJn(A)
Case Else
    Stop
End Select
Evl_24129_SqlPhrase = O
End Function

Private Function Evl_241291__SqlSel(A As KPD) As StrOpt
Dim O$
    With Evl_2412A_SqlSelFldLis(A)
        If Not .Som Then Exit Function
        O = .Str
    End With
Evl_241291__SqlSel = SomStr("Select|" & O)
End Function

Private Function Evl_241292__SqlInto(K$) As StrOpt
Debug.Print "It should never be called, because, SqlKw-Into is no need to code."
Debug.Print "In $ or @"
Stop
Dim T$
    With BrkRev(K$, ".")
        If .S2 <> "Into" Then Stop
        T = TakAftRev(.S1, ".")
    End With
Evl_241292__SqlInto = SomStr("|   Into " & T)
End Function

Private Function Evl_241293__SqlFm(L3Prm$) As StrOpt
Evl_241293__SqlFm = SomStr("|  From " & L3Prm)
End Function

Private Function Evl_241294__SqlAnd(A As KPD) As StrOpt
Dim Sy$()
    With Evl_PrmLvs(A)
        If Not .Som Then
            Exit Function
        End If
        Sy = .Sy
    End With
Evl_241294__SqlAnd = SomStr(Join(AyAddPfx("|    And ", Sy)))
End Function

Private Function Evl_241295__SqlGp(A As KPD) As StrOpt
Dim B As KPD
    B = A
    If Not IsSfx(B.K, ".Gp") Then Stop
    B.K = RmvSfx(B.K, ".Gp") & ".Sel"
Dim Sy$()
    With Evl_PrmLvs(B)
        If Not .Som Then
            Exit Function
        End If
        Sy = AyRmvEmpty(.Sy)
    End With
Dim O$
    Sy = AyAlignL(Sy)
    Sy = AyAddPfx(Sy, "|    ")
    O = JnComma(Sy)
    O = "|  Group By" & O
    Evl_241295__SqlGp = SomStr(O)
End Function

Private Function Evl_241296__SqlUpd(K$) As StrOpt
Debug.Print "It should never be called, because, SqlKw-Upd is no need to code."
Debug.Print "In $ or @"
Stop
Dim Ns$: Ns = Key_Ns(K)
Dim A$: A = TakAftRev(Ns, ".")
Dim T$: T = TakBefRev(A, "#")
Dim O$
    O = "Update #" & T
Evl_241296__SqlUpd = SomStr(O)
End Function

Private Function Evl_241297__SqlSet(A As KPD) As StrOpt
Dim Ay$()
    Ay = SplitLvs(A.P)
Dim Sy$()
    With Evl_PrmLvs(A)
        If Not .Som Then Exit Function
        Sy = .Sy
    End With
Dim O$()
    Dim J%, B$
    For J = 0 To UB(Ay)
        B = Ay(J) & " = " & Sy(J)
        Push O, B
    Next
Evl_241297__SqlSet = SomStr("|  Set|    " & Join(O, ",|    "))
End Function

Private Function Evl_241298__SqlSelDis(A As KPD) As StrOpt
Dim Ay$()
    Ay = SplitLvs(A.P)
Dim Sy$()
    With Evl_PrmLvs(A)
        If Not .Som Then Exit Function
        Sy = .Sy
    End With
Dim O$()
    Dim J%, B$
    For J = 0 To UB(Ay)
        B = Sy(J) & " " & Ay(J)
        Push O, B
    Next
Evl_241298__SqlSelDis = SomStr("Select|  Distinct|    " & Join(O, ",|    "))
End Function

Private Function Evl_241299__SqlWh(A As KPD) As StrOpt
Dim O$
    O = "|  Where " & FmtMacroDic(A.P, A.D)
Evl_241299__SqlWh = SomStr(O)
End Function

Private Function Evl_24129A__SqlJn(A As KPD) As StrOpt
Dim Ay$()
    Ay = SplitLvs(A.P)
Dim Sy$()
    With Evl_PrmLvs(A)
        If Not .Som Then
            Exit Function
        End If
        Sy = .Sy
    End With
Dim O$()
    O = AyAddPfx("|    ", AyRmvEmpty(Sy))
Evl_24129A__SqlJn = SomStr(Join(O))
End Function

Private Function Evl_2412A_SqlSelFldLis(A As KPD) As StrOpt

End Function

Private Function Evl_PrmAy(K$, PrmAy$(), Dic As Dictionary) As SyOpt
'{Prm} in PrmAy is either with [.] or not.
'If FstChr is [.], just the {Prm}
'If with [.], just use {Prm} to lookup value in Dic
'If no [.], use {K}.{Prm} to lookup value in Dic
Dim Vy$(), PrmTerm
    For Each PrmTerm In PrmAy
        With Evl_PrmAy_1__EvlOneTerm(K, PrmTerm, Dic)
            If Not .Som Then Exit Function
            Push Vy, .Str
        End With
    Next
Evl_PrmAy = SomSy(Vy)
End Function

Private Function Evl_PrmAy_1__EvlOneTerm(K$, PrmTerm, Dic As Dictionary) As StrOpt
If Not Evl_PrmAy_11_IsTermValid(K, PrmTerm, Dic) Then Exit Function

If FstChr(PrmTerm) = "." Then
    Evl_PrmAy_1__EvlOneTerm = SomStr(PrmTerm)
    Exit Function
End If
If HasSubStr(PrmTerm, ".") Then
    Evl_PrmAy_1__EvlOneTerm = SomStr(Dic(PrmTerm))
    Exit Function
End If
Evl_PrmAy_1__EvlOneTerm = SomStr(Dic(K & "." & PrmTerm))
End Function

Private Function Evl_PrmAy_11_IsTermValid(K$, PrmTerm, Dic As Dictionary) As Boolean
Evl_PrmAy_11_IsTermValid = True
If FstChr(PrmTerm) = "." Then Exit Function
If HasSubStr(PrmTerm, ".") Then
    If Dic.Exists(PrmTerm) Then Exit Function
End If
If Dic.Exists(K & "." & PrmTerm) Then Exit Function
Evl_PrmAy_11_IsTermValid = False
End Function

Private Function Evl_PrmLvs(A As KPD) As SyOpt
Dim Ay$(): Ay = SplitLvs(A.P)
Evl_PrmLvs = Evl_PrmAy(A.K, Ay, A.D)
End Function

Private Function Evl_SetFixStr(A() As L12SOPS) As L12SOPS()
Dim O() As L12SOPS: O = A
Dim J&
For J = 0 To L12SOPS_UB(O)
    If O(J).Op = eFixStr Then
        O(J).Don = True
        O(J).S = RplFstChrToIf(O(J).Prm, ".", " ")
    End If
Next
Evl_SetFixStr = O
End Function

Private Function Evl_ToDic(A() As L12SOPS) As Dictionary
Dim PrmQ As Dictionary: Set PrmQ = Evl_ToDic_1_FmPrmQ(A)
Dim SngD As Dictionary: Set SngD = Evl_ToDic_2_FmSngNm(A)
Dim MulD As Dictionary: Set MulD = Evl_ToDic_3_FmMulNm(A)
Set Evl_ToDic = DicAdd(PrmQ, SngD, MulD)
End Function

Private Function Evl_ToDic_1_FmPrmQ(A() As L12SOPS) As Dictionary
Dim J%, O As New Dictionary, K$, V As Boolean
For J = 0 To L12SOPS_UB(A)
    If L12SOPS_IsPrmSwitch(A(J)) Then
        K = A(J).L2
        Select Case A(J).Prm
        Case "1": V = True
        Case "0": V = False
        Case Else: Stop
        End Select
        O.Add K, V
    End If
Next
Set Evl_ToDic_1_FmPrmQ = O
End Function

Private Function Evl_ToDic_2_FmSngNm(A() As L12SOPS) As Dictionary
Dim B() As L12SOPS: B = L12SOPSAy_SelSngNm(A)
Dim J%, O As New Dictionary
For J = 0 To L12SOPS_UB(B)
    If B(J).Don Then
        O.Add L12SOPS_Key(B(J)), B(J).S
    End If
Next
Set Evl_ToDic_2_FmSngNm = O
End Function

Private Function Evl_ToDic_3_FmMulNm(A() As L12SOPS) As Dictionary
Dim B() As L12SOPS
    B = A
Dim O As New Dictionary
Dim X%
For X = 0 To 99
    With L12SOPSAy_MulNmS1S2Opt(B)
        If Not .Som Then
            Set Evl_ToDic_3_FmMulNm = O
            Exit Function
        End If
        With .S1S2
            O.Add .S1, .S2
            B = L12SOPSAy_MinusL1L2(B, .S1, .S2)
        End With
    End With
Next
Never
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

Private Sub KPD_Brw(A As KPD)
AyBrw KPD_Ly(A)
End Sub

Private Function KPD_Ly(A As KPD)
Dim O As New Dictionary
    O.Add ".Key", A.K
    O.Add ".Prm", A.P
    Set O = DicAdd(O, A.D)
KPD_Ly = DicLy(O)
End Function

Private Sub KPDDmp(A1 As KPD)
AyDmp KPDLy(A1)
End Sub

Private Function KPDLy(A As KPD) As String()
Dim D As Dictionary
    Set D = DicClone(A.D)
    D.Add "**Key", A.K
    D.Add "**Prm", A.P
    Set D = DicSrt(D)
KPDLy = DicLy(D)
End Function

Private Function L123(LinI%, L1$, L2$, L3$) As L123
Dim O As L123
With O
    .L1 = L1
    .L2 = L2
    .L3 = L3
    .LinI = LinI
End With
L123 = O
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

Private Function L123_IsInDic(A As L123, D As Dictionary) As Boolean
Dim K$: K = L123_Key(A)
L123_IsInDic = D.Exists(K)
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

Private Function L123_L12SOPS(A As L123) As L12SOPS
Dim O As L12SOPS
Dim B As L3
B = L3Str_Brk(A.L3)
With O
    .LinI = A.LinI
    .L1 = A.L1
    .L2 = A.L2
    .Prm = B.Prm
    .Switch = B.Switch
    .Op = B.Op
End With
L123_L12SOPS = O
End Function

Private Sub L123_Push(Ay() As L123, M As L123)
Dim N%: N = L123_Sz(Ay)
ReDim Preserve Ay(N)
Ay(N) = M
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
    O.L3 = L3Str_Brk(A.L3)
End With
L123_WrkDr = O
End Function

Private Sub L123Ay_Brw(A() As L123, Optional Fnn)
DrsBrw L123Ay_Drs(A), Fnn:=Fnn
End Sub

Private Function L123Ay_DistOpSy(A() As L123) As String()
Dim J%, O() As eOp
For J = 0 To L123_UB(A)
    Push O, L3Str_Brk(A(J).L3).Op
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

Private Function L123Ay_Dt(A() As L123) As Dt
Dim O As Dt
With L123Ay_Drs(A)
    O.Dry = .Dry
    O.Fny = .Fny
End With
O.DtNm = "L123"
L123Ay_Dt = O
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

Private Function L123Ay_Ly(A() As L123) As String()
L123Ay_Ly = DrsLy(L123Ay_Drs(A))
End Function

Private Function L123Ay_MinusDic(A() As L123, D As Dictionary) As L123()
Dim O() As L123, J%
For J = 0 To L123_UB(A)
    If Not L123_IsInDic(A(J), D) Then
        L123_Push O, A(J)
    End If
Next
L123Ay_MinusDic = O
End Function

Private Function L123Ay_PrmDic(A() As L123) As Dictionary
Dim O As New Dictionary
    Dim B() As L123: B = L123Ay_SelPrm(A)
    Dim J%
    Dim K$, V$
    For J = 0 To L123_UB(B)
        With B(J)
            V = L3Str_Brk(.L3).Prm
            K = .L1 & "." & .L2
            O.Add K, V
        End With
    Next
Set L123Ay_PrmDic = O
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

Private Function L123Ay_SelNsPfxLvs(L123Ay() As L123, NsPfxLvs$) As L123()
Dim J%, NsPfxAy$(), M As L123, O() As L123
NsPfxAy = SplitLvs(NsPfxLvs)
For J = 0 To L123_UB(L123Ay)
    M = L123Ay(J)
    If PfxAyHas(NsPfxAy, M.L1) Then
        L123_Push O, M
    End If
Next
L123Ay_SelNsPfxLvs = O
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
Dim O As Dictionary
    Set O = SwitchPrmDic(PrmDic)    '<== Any Prm with Nm = ?XXX, promote them as Switch
    Dim Dic As Dictionary
    Set Dic = DicAdd(PrmDic, O)
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

Private Function L12SOPS_Dr(A As L12SOPS) As Variant()
With A
    L12SOPS_Dr = Array(.LinI, .L1, .L2, L12SOPS_Key(A), .Switch, OpStr(.Op), .Prm, .S, .Don)
End With
End Function

Private Function L12SOPS_IsInDic(A As L12SOPS, D As Dictionary) As Boolean
Dim K$: K = L12SOPS_Key(A)
L12SOPS_IsInDic = D.Exists(K)
End Function

Private Function L12SOPS_IsPrmSwitch(A As L12SOPS) As Boolean
With A
    If .L1 = "Prm" Then
        If IsPfx(.L2, "?") Then
            L12SOPS_IsPrmSwitch = True
        End If
    End If
End With
End Function

Private Function L12SOPS_Key$(A As L12SOPS)
If A.L1 = "?" Then
    L12SOPS_Key = "?" & A.L2
Else
    L12SOPS_Key = A.L1 & "." & A.L2
End If
End Function

Private Sub L12SOPS_Push(A() As L12SOPS, M As L12SOPS)
Dim N%: N = L12SOPS_Sz(A)
ReDim Preserve A(N)
A(N) = M
End Sub

Private Function L12SOPS_Sz%(A() As L12SOPS)
On Error Resume Next
L12SOPS_Sz = UBound(A) + 1
End Function

Private Function L12SOPS_UB%(A() As L12SOPS)
L12SOPS_UB = L12SOPS_Sz(A) - 1
End Function

Private Function L12SOPSAy_Brw(A() As L12SOPS, Optional Fnn)
DtBrw L12SOPSAy_Dt(A), Fnn
End Function

Private Function L12SOPSAy_Brw_NoL1L2(A() As L12SOPS, Optional Fnn)
DtBrw L12SOPSAy_Dt_NoL1L2(A), Fnn
End Function

Private Function L12SOPSAy_DonCnt%(A() As L12SOPS)
Dim O%
Dim J%
For J = 0 To L12SOPS_UB(A)
    If A(J).Don Then O = O + 1
Next
L12SOPSAy_DonCnt% = O
End Function

Private Function L12SOPSAy_Drs(A() As L12SOPS) As Drs
L12SOPSAy_Drs = Drs(L12SOPSAy_Fny, L12SOPSAy_Dry(A))
End Function

Private Function L12SOPSAy_Dry(A() As L12SOPS) As Variant()
Dim Dr, O(), J%
For J = 0 To L12SOPS_UB(A)
    Push O, L12SOPS_Dr(A(J))
Next
L12SOPSAy_Dry = O
End Function

Private Function L12SOPSAy_Dt(A() As L12SOPS) As Dt
L12SOPSAy_Dt = Dt("L12SOPS", L12SOPSAy_Drs(A))
End Function

Private Function L12SOPSAy_Dt_NoL1L2(A() As L12SOPS) As Dt
Dim B As Dt: B = L12SOPSAy_Dt(A)
L12SOPSAy_Dt_NoL1L2 = DtDrpCol(B, "L1 L2")
End Function

Private Function L12SOPSAy_Fny() As String()
L12SOPSAy_Fny = SplitLvs("LinI L1 L2 Key Switch OpStr Prm S Done")
End Function

Private Function L12SOPSAy_Has(A() As L12SOPS, I As L12SOPS) As Boolean
Dim J%
With I
    For J = 0 To L12SOPS_UB(A)
        If A(J).L1 <> I.L1 Then GoTo Nxt
        If A(J).L2 <> I.L2 Then GoTo Nxt
        L12SOPSAy_Has = True
        Exit Function
Nxt:
    Next
End With
End Function

Private Function L12SOPSAy_IsAllDon(A() As L12SOPS) As Boolean
L12SOPSAy_IsAllDon = L12SOPSAy_NotDonCnt(A) = 0
End Function

Private Function L12SOPSAy_IsAllSwitchSet(A() As L12SOPS) As Boolean
Dim J%
For J = 0 To L12SOPS_UB(A)
    With A(J)
        If .L1 = "?" Then
            If Not .Don Then Exit Function
        End If
    End With
Next
L12SOPSAy_IsAllSwitchSet = True
End Function

Private Function L12SOPSAy_IsEmpty(A() As L12SOPS) As Boolean
L12SOPSAy_IsEmpty = L12SOPS_Sz(A) = 0
End Function

Private Function L12SOPSAy_Ky(A() As L12SOPS) As String()
Dim O$(), J%
For J = 0 To L12SOPS_UB(A)
    Push O, L12SOPS_Key(A(J))
Next
L12SOPSAy_Ky = O
End Function

Private Function L12SOPSAy_Ky_Uniq(A() As L12SOPS) As String()
L12SOPSAy_Ky_Uniq = AyDist(L12SOPSAy_Ky(A))
End Function

Private Function L12SOPSAy_Ky_UniqSrt(A() As L12SOPS) As String()
L12SOPSAy_Ky_UniqSrt = L12SOPSAy_Ky_Uniq(A)
End Function

Private Function L12SOPSAy_MinusL1L2(A() As L12SOPS, L1$, L2$) As L12SOPS()

End Function

Private Function L12SOPSAy_MulNmS1S2Opt(A() As L12SOPS) As S1S2Opt

End Function

Private Function L12SOPSAy_NotDonCnt%(A() As L12SOPS)
Dim J%, O%
For J = 0 To L12SOPS_UB(A)
    If Not A(J).Don Then O = O + 1
Next
L12SOPSAy_NotDonCnt = O
End Function

Private Function L12SOPSAy_SelSngNm(A() As L12SOPS) As L12SOPS()
Dim Ky$(): Ky = L12SOPSAy_SngKy(A)
Dim O() As L12SOPS, J%, K$
For J = 0 To L12SOPS_UB(A)
    K = L12SOPS_Key(A(J))
    If AyHas(Ky, K) Then
        L12SOPS_Push O, A(J)
    End If
Next
L12SOPSAy_SelSngNm = O
End Function

Private Function L12SOPSAy_SngKy(A() As L12SOPS) As String()
Dim Ky$(): Ky = L12SOPSAy_Ky(A)
L12SOPSAy_SngKy = AySelSngEle(Ky)
End Function

Private Function L12SOPSAy_Srt(A() As L12SOPS) As L12SOPS()
Dim I&(): I = L12SOPSAy_SrtIntoIdxAy(A)
Dim O() As L12SOPS: O = A
Dim J%
For J = 0 To L12SOPS_UB(A)
    O(J) = A(I(J))
Next
L12SOPSAy_Srt = O
End Function

Private Function L12SOPSAy_SrtIntoIdxAy(A() As L12SOPS) As Long()
Dim K$(), J%
For J = 0 To L12SOPS_UB(A)
    Push K, L12SOPS_Key(A(J))
Next
L12SOPSAy_SrtIntoIdxAy = AySrtIntoIdxAy(K)
End Function

Private Function L12SOPSStr$(A As L12SOPS)
Dim K$, S$, P$, Op$, L$
    K = " " & L12SOPS_Key(A)
    With A
        S = .Switch: If S <> "" Then S = " " & S
        P = .Prm
        Op = " " & OpStr(.Op)
        L = .LinI
    End With
L12SOPSStr = L & K & S & Op & P
End Function

Private Sub L1TrmLin_Brk(L1TrmLin, OL1$, OL2$, OL3$)
'L1TrmLin is either have only 1 term for L1 or 3 term for L1,2,3
If HasSubStr(L1TrmLin, " ") Then
    Dim L$: L = L1TrmLin
    OL1 = ParseTerm(L): If OL1 = "" Then Stop
    OL2 = ParseTerm(L): If OL2 = "" Then Stop
    OL3 = L: If OL3 = "" Then Stop
Else
    OL1 = L1TrmLin
End If
End Sub

Private Function L3_Key$(A As L123)
L3_Key = A.L1 & "." & A.L2
End Function

Private Function L3Str_Brk(L3Str) As L3
Dim L$: L = Trim(L3Str)
If L3Str = "" Then Exit Function
Dim Switch$
    If FstChr(L) = "?" Then
        Switch = FstTerm(L)
        L = RmvFstTerm(L)
    End If
Dim OpStr$
    OpStr = FstTerm(L)
   
Dim O As L3
With O
    .L3 = L3Str
    .Switch = Switch
    .OpStr = OpStr
    .Op = Op(OpStr)
    .Prm = RestTerm(L)
End With
L3Str_Brk = O
End Function

Private Function L3Str_SwitchVal(L3Str$, Dic As Dictionary) As Boolean
Dim A As L3: A = L3Str_Brk(L3Str)
L3Str_SwitchVal = Evl_23___EvlSwitchVal(A.Op, A.Prm, Dic)
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
Case "*": O = eSqlStmt
Case "$In": O = eExpIn
Case "#": O = eSqlPhrase
Case "#Sel": O = eSqlSel
Case "$": O = eMac
Case ".": O = eFixStr
Case ".And": O = eFixAnd
Case ".Or": O = eFixOr
Case ".Comma": O = eFixComma
Case ".Drp": O = eFixDrp
Case ".Eq": O = eFixEq
Case ".Ne": O = eFixNe
Case ".Or": O = eFixOr
Case "@": O = eExp
Case "@Comma": O = eExpComma
Case Else: Stop: O = eUnknown
End Select
Op = O
End Function

Private Function Op_AlwSwitchOpAy() As eOp()
Static X As Boolean, O
Dim I
If Not X Then
    Dim A() As eOp
    O = A
    X = True
    For Each I In Array(eOp.eMac, eOp.eExp, eOp.eExpComma, _
        eOp.eSqlStmt, eOp.eFixStr, eOp.eSqlSel, _
        eOp.eExpIn)
        Push O, I
    Next
End If
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
    eOp.eExpIn, _
    eOp.eExpComma, _
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

Private Function OpStr$(A As eOp)
Dim O$
Select Case A
Case eExpIn: O = "$In"
Case eRun: O = "!"
Case eSqlPhrase: O = "#"
Case eSqlSel:    O = "#Sel"
Case eFixEq:     O = ".Eq"
Case eSqlStmt:   O = "*"
Case eExpComma:  O = "@Comma"
Case eExp:       O = "@"
Case eFixAnd:    O = ".And"
Case eFixOr:    O = ".Or"
Case eFixEq:    O = ".Eq"
Case eFixNe:    O = ".Ne"
Case eFixDrp:    O = ".Drp"
Case eFixComma:  O = ".Comma"
Case eFixStr: O = "."
Case eMac:   O = "$"
Case Else: Stop: O = "?Unknown"
End Select
OpStr = O
End Function

Private Function PrmIsLisOfRf_Of_OpAy() As eOp()
Static O() As eOp, X As Boolean
If Not X Then
    X = True
    Push O, eOp.eSqlPhrase
    Push O, eOp.eExp
    Push O, eOp.eMac
End If
PrmIsLisOfRf_Of_OpAy = O
End Function

Private Function PrmIsLisOfRf_Of_SqlKw() As String()
Static X As Boolean, O
If Not X Then
    Dim A$()
    O = A
    X = True
    Push O, "Jn"
    Push O, "And"
    Push O, "Or"
    Push O, "Gp"
    Push O, "Sel"
    Push O, "SelDis"
End If
PrmIsLisOfRf_Of_SqlKw = O
End Function

Private Function SomErDr(ErDr) As ErDrOpt
With SomErDr
    .Som = True
    .ErDr = ErDr
End With
End Function

Private Function Soml123(LinI%, L1$, L2$, L3$) As L123Opt
Soml123.Som = True
Soml123.L123 = L123(LinI%, L1, L2, L3)
End Function

Private Function SomL12SOPSAy(A() As L12SOPS) As L12SOPSAyOpt
With SomL12SOPSAy
    .Som = True
    .L12SOPSAy = A
End With
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
Dim Ly1$(): Ly1 = Sql3Ly_ValidatedLy(Ly)
If AyIsEmpty(Ly1) Then
    Exit Function
End If
If AyIsEq(Ly, Ly1) Then Exit Function
AyWrt Ly1, Ft
Sql3Ft_WrtEr = True
End Function

Private Function Sql3Ly_AddEr(Sql3Ly$(), ErDry()) As String()
Dim W%, O$()
    O = Sql3Ly_No3DashLy(Sql3Ly)
    O = AyRTrim(O)
    W = AyWdt(O)
Dim LinI%, Dr
DryBrw ErDry
Dim ErDry1()
    ErDry1 = DryMge(ErDry, 1, " | ")
'DryBrw ErDry1: Stop
For Each Dr In ErDry1
    LinI = Dr(0)
    If Len(O(LinI)) > W Then Stop
    O(LinI) = AlignL(O(LinI), W) & " --- " & Dr(1)
Next
AyRmvEmptyEleAtEnd O
Push O, FmtQQ("--- [?] error(s)", Sz(ErDry))
Sql3Ly_AddEr = O
End Function

Private Function Sql3Ly_ErDry(Sql3Ly$()) As Variant()
Sql3Ly_ErDry = Vdt(Sql3Ly_Wy(Sql3Ly))
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
Sql3Ly_No3DashLy = AyMapIntoSy(Sql3Ly, "Rmv3Dash")
End Function

Private Function Sql3Ly_TrmLy(Sql3Ly$()) As String()
Sql3Ly_TrmLy = AyMapIntoSy(Sql3Ly, "Rmv2Dash")
End Function

Private Function Sql3Ly_ValidatedLy(Sql3Ly$()) As String()
Dim ErDry(): ErDry = Sql3Ly_ErDry(Sql3Ly)
If AyIsEmpty(ErDry) Then Exit Function
Sql3Ly_ValidatedLy = Sql3Ly_AddEr(Sql3Ly, ErDry)
End Function

Private Function Sql3Ly_Wy(Sql3Ly$()) As WrkDr()
Dim A() As L123: A = Evl_1___Pass1(Sql3Ly)
Dim B() As L123: B = Evl_1___Pass2(A)
Sql3Ly_Wy = L123Ay_Wy(B)
End Function

Private Sub SwitchNm_Assert(SwitchNm$)
If FstChr(SwitchNm) <> "?" Then Stop
End Sub

Private Function SwitchPrm_BoolAy(SwitchPrm$, Dic As Dictionary) As Boolean()
Dim TermAy$()
    TermAy = SplitSpc(SwitchPrm)
Dim O() As Boolean
    ReDim O(UB(TermAy))
    Dim J%
    For J = 0 To UB(TermAy)
        With SwitchTerm_Val(TermAy(J), Dic)
            If Not .Som Then
'                Stop
                Exit Function
            End If
            O(J) = .Bool
        End With
    Next
SwitchPrm_BoolAy = O
End Function

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

Private Function SwitchPrmDic(PrmDic As Dictionary) As Dictionary
'Prm may have Nm=?xxx.  These are the SwitchPrm.  It must have value 1 or 0
'Return a Dic of this SwitchPrm
If DicIsEmpty(PrmDic) Then Set SwitchPrmDic = New Dictionary: Exit Function
Dim K, V, O As New Dictionary
For Each K In PrmDic.Keys
    If IsPfx(K, "Prm.?") Then
        V = PrmDic(K)
        If V <> "1" And V <> "0" Then Stop
        O.Add RmvPfx(K, "Prm."), V = "1"
    End If
Next
Set SwitchPrmDic = O
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

Private Function SwitchWrkDr_Val(A As WrkDr, Dic As Dictionary) As Boolean
Select Case A.L3.Op
Case eFixEq:  SwitchWrkDr_Val = Evl_231__EvlSwitchVal_ForEq(A.L3.Prm, Dic)
Case eFixNe:  SwitchWrkDr_Val = Evl_232__EvlSwitchVal_ForNe(A.L3.Prm, Dic)
Case eFixOr:  SwitchWrkDr_Val = Evl_233__EvlSwitchVal_ForOr(A.L3.Prm, Dic)
Case eFixAnd: SwitchWrkDr_Val = Evl_234__EvlSwitchVal_ForAnd(A.L3.Prm, Dic)
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
    Case 1: L1TrmLin_Brk TrmLin, L1, L2, L3              ' Assume L1 can only have one term or all 3 terms
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

Private Function Vdt(Wy() As WrkDr) As Variant()
Dim O()
PushAy O, Vdt_InvalidOp(Wy)
PushAy O, Vdt_GpDotXXX_IsNoNeed(Wy)
PushAy O, Vdt_MissingRf(Wy)   ' Som Operations, the parameter is a list of PrmTerm.  Each Term need to be refer further defintion.  It is to check such definition exist
PushAy O, Vdt_NotAlwSwitch(Wy)
PushAy O, Vdt_Switch(Wy)
PushAy O, Vdt_NoPrm(Wy)
PushAy O, Vdt_NoSql(Wy)
Vdt = O
End Function

Private Function Vdt_GpDotXXX_IsNoNeed(Wy() As WrkDr) As Variant()
Dim O()
Dim J%
For J = 0 To Wy_UB(Wy)
    If IsSfx(Wy(J).Ns, ".Gp") Then
        Push O, Array(Wy(J).LinI, "No need to define Itm-[XXX.Gp.FldNm], because such definition will use Itm-[XXX.Sel.FldNm]")
    End If
Next
Vdt_GpDotXXX_IsNoNeed = O
End Function

Private Function Vdt_InvalidOp(Wy() As WrkDr) As Variant()
Dim J%
Dim O()
For J = 0 To UBound(Wy)
    With Wy(J).L3
        If .Op = eOp.eUnknown Then
            Push O, Array(Wy(J).LinI, FmtQQ("Invalid Op[?]", .OpStr))
        End If
    End With
Next
Vdt_InvalidOp = O
End Function

Private Function Vdt_MissingRf(Wy() As WrkDr) As Variant()
Dim W() As WrkDr
    W = Vdt_MissingRf_Wy(Wy)
Dim O()
Dim J%, KeyDic As Dictionary
Set KeyDic = Wy_KeyDic(Wy)
For J = 0 To Wy_UB(W)
    With Vdt_MissingRf_ErDrOpt(W(J), KeyDic)
        If .Som Then
            Push O, .ErDr
        End If
    End With
Next
Vdt_MissingRf = O
End Function

Private Function Vdt_MissingRf_ErDrOpt(A As WrkDr, NsNmDic As Dictionary) As ErDrOpt
Dim Ay$(): Ay = SplitLvs(A.L3.Prm)
Dim O$(), K$
If A.Nm = "Gp" Then
    K = A.Ns & ".Sel"
Else
    K = WrkDr_Key(A)
End If
Dim Term
For Each Term In Ay
    If Vdt_MissingRf_IsTerm_MissingRf(K, Term, NsNmDic) Then
        Push O, Term
    End If
Next
If AyIsEmpty(O) Then Exit Function
Dim Er
    Dim B$
    B = "Invalid term: " & Join(AyQuote(O, "[]"), " ")
    Er = Array(A.LinI, B)
Vdt_MissingRf_ErDrOpt = SomErDr(Er)
End Function

Private Function Vdt_MissingRf_IsTerm_MissingRf(K$, Term, NsNmDic As Dictionary) As Boolean
'[T] can be:
'1. {K}.{Term}, or,   if {Term} does not contains [.]
'2 {Term}     , or,   if fstchr-{Term} is not [.] and {Term} contains [.]
'3 {Term}    ,       if fstchr-{Term} is [.]
'For case-1 and 2, return true if [T] is NOT in {NsNmDic} else return false
'For case-3,       return false
If FstChr(Term) = "." Then
    Exit Function
End If
Dim T$
    If HasSubStr(Term, ".") Then
        T = Term
    Else
        T = K & "." & Term
    End If
Vdt_MissingRf_IsTerm_MissingRf = Not NsNmDic.Exists(T)
End Function

Private Function Vdt_MissingRf_IsWrkDrHasRfPrm(A As WrkDr) As Boolean
Vdt_MissingRf_IsWrkDrHasRfPrm = True
If A.L3.Op = eSqlPhrase Then
    If AyHas(PrmIsLisOfRf_Of_SqlKw, A.Nm) Then Exit Function
End If
If AyHas(PrmIsLisOfRf_Of_OpAy, A.L3.Op) Then Exit Function
Vdt_MissingRf_IsWrkDrHasRfPrm = False
End Function

Private Function Vdt_MissingRf_Wy(Wy() As WrkDr) As WrkDr()
Dim O() As WrkDr
Dim J%
For J = 0 To Wy_UB(Wy)
    If Vdt_MissingRf_IsWrkDrHasRfPrm(Wy(J)) Then
        Wy_Push O, Wy(J)
    End If
Next
Vdt_MissingRf_Wy = O
End Function

Private Function Vdt_NoPrm(Wy() As WrkDr) As Variant()
Dim J%
For J = 0 To Wy_UB(Wy)
    If Wy(J).Ns = "Prm" Then Exit Function
Next
Vdt_NoPrm = Array(Array(0, "Warning: No Prml namespace"))
End Function

Private Function Vdt_NoSql(Wy() As WrkDr) As Variant()
Dim J%
For J = 0 To Wy_UB(Wy)
    If Wy(J).Ns = "Sql" Then Exit Function
Next
Vdt_NoSql = Array(Array(0, "Warning: No Sql namespace"))
End Function

Private Function Vdt_NotAlwSwitch(Wy() As WrkDr) As Variant()
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
Vdt_NotAlwSwitch = O
End Function

Private Function Vdt_Switch(Wy() As WrkDr) As Variant()
'Return ErDry if any Wy(..) has switch not defined in Wy()
Dim J%, O(), D As Dictionary
Set D = AyDic(Wy_SwitchNy(Wy))
For J = 0 To UBound(Wy)
    With Wy(J).L3
        If .Switch = "" Then GoTo Nxt
        If D.Exists(.Switch) Then GoTo Nxt
        Push O, Array(Wy(J).LinI, FmtQQ("Switch [?] not exist", .Switch))
    End With
Nxt:
Next
Vdt_Switch = O
End Function

Private Function WrkDr_HasSwitch(A As WrkDr) As Boolean
WrkDr_HasSwitch = A.L3.Switch <> ""
End Function

Private Function WrkDr_Key$(A As WrkDr)
WrkDr_Key = A.Ns & "." & A.Nm
End Function

Private Sub Wy_Brw(Wy() As WrkDr, Optional SwitchDic As Dictionary)
DrsBrw Wy_Drs(Wy, SwitchDic)
End Sub

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

Private Function Wy_Fny() As String()
Wy_Fny = SplitSpc("LinI Ns Nm Switch SwitchVal Op Prm L3")
End Function

Private Function Wy_Has(A() As WrkDr, B As WrkDr) As Boolean
Dim J%
For J = 0 To Wy_UB(A)
    If A(J).LinI = B.LinI Then Wy_Has = True: Exit Function
Next
End Function

Private Function Wy_HasSwitchDefined(Wy() As WrkDr, Switch$) As Boolean
Debug.Print "Wy_HasSwitchDefined: ", Switch
Dim J%
For J = 0 To UBound(Wy)
    With Wy(J).L3
        If .Switch = "" Then GoTo Nxt
        If .Switch <> Switch Then
            Wy_HasSwitchDefined = True
            Exit Function
        End If
    End With
Nxt:
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

Private Function Wy_KeyDic(Wy() As WrkDr) As Dictionary
Dim O As New Dictionary
Dim J%, K$
For J = 0 To Wy_UB(Wy)
    K = WrkDr_Key(Wy(J))
    If Not O.Exists(K) Then
        O.Add K, True
    End If
Next
Set Wy_KeyDic = O
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

Private Function Wy_Ny(Wy() As WrkDr) As String()
Dim O$()
Dim J%
For J = 0 To Wy_UB(Wy)
    If Wy(J).Ns = "?" Then
        Push O, "?" & Wy(J).Nm
    ElseIf Wy(J).Ns = "Prm" Then
        If IsPfx(Wy(J).Nm, "?") Then
            Push O, Wy(J).Nm
        End If
    End If
Next
Wy_Ny = O
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

Private Function Wy_RmvItms(A() As WrkDr, IdxAy%()) As WrkDr()
If AyIsEmpty(IdxAy) Then Wy_RmvItms = A: Exit Function
Dim O() As WrkDr, J%
For J = 0 To Wy_UB(A)
    If Not AyHas(IdxAy, J) Then Wy_Push O, A(J)
Next
Wy_RmvItms = O
End Function

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

Private Function Wy_SelWithSwith(A() As WrkDr) As WrkDr()
Dim O() As WrkDr, J%
For J = 0 To Wy_UB(A)
    If WrkDr_HasSwitch(A(J)) Then
        Wy_Push O, A(J)
    End If
Next
Wy_SelWithSwith = O
End Function

Private Function Wy_SwitchNy(A() As WrkDr) As String()
Dim A1$(): A1 = Wy_SwitchNy_FmPrmNsWithNmBegWithQuestionMrk(A)
Dim A2$(): A2 = Wy_SwitchNy_FmNsIsQuestionMrk(A)
Wy_SwitchNy = AyAdd(A1, A2)
End Function

Private Function Wy_SwitchNy_FmNsIsQuestionMrk(A() As WrkDr) As String()
Dim O$(), J%
For J = 0 To Wy_UB(A)
    If A(J).Ns = "?" Then
        Push O, "?" & A(J).Nm
    End If
Next
Wy_SwitchNy_FmNsIsQuestionMrk = O
End Function

Private Function Wy_SwitchNy_FmPrmNsWithNmBegWithQuestionMrk(A() As WrkDr) As String()
Dim O$(), J%
For J = 0 To Wy_UB(A)
    If A(J).Ns = "Prm" Then
        If IsPfx(A(J).Nm, "?") Then
            Push O, A(J).Nm
        End If
    End If
Next
Wy_SwitchNy_FmPrmNsWithNmBegWithQuestionMrk = O
End Function

Private Function Wy_Sz&(Wy() As WrkDr)
On Error Resume Next
Wy_Sz = UBound(Wy) + 1
End Function

Private Function Wy_UB&(Wy() As WrkDr)
Wy_UB = Wy_Sz(Wy) - 1
End Function

Private Function Wy_Ws(Wy() As WrkDr, Optional SwitchDic As Dictionary) As Worksheet
DrsWs Wy_Drs(Wy, SwitchDic)
End Function

Private Function ZZEvl_1____FmSql3Ly_ToL123Ay() As L123()
ZZEvl_1____FmSql3Ly_ToL123Ay = Evl_1____FmSql3Ly_ToL123Ay(ZZSql3Ly)
End Function

Private Function ZZL12SOPSAy(Optional Sql3FtVarNm = "") As L12SOPS()
Dim A() As L123: A = Evl_1___Pass1(ZZSql3Ly(Sql3FtVarNm))
ZZL12SOPSAy = Evl_1___Pass2(A)
End Function

Private Function ZZSql3Ft$(Optional Sql3FtVarNm = "")
Dim V$
If Not IsEmpty(Sql3FtVarNm) Then V = "-" & Sql3FtVarNm
ZZSql3Ft = TstResPth & FmtQQ("SalRpt?.Sql3", V)
End Function

Private Sub ZZSql3Ft_Edt(Optional Sql3FtVarNm = "")
FtBrw ZZSql3Ft(Sql3FtVarNm)
End Sub

Private Sub ZZSql3Ft_Fix()
Dim O$(): O = ZZSql3Ly
Dim J%
For J = 0 To UB(O)
    O(J) = Replace(O(J), Chr(160), " ")
Next
AyWrt O, ZZSql3Ft
End Sub

Private Function ZZSql3Ly(Optional Sql3FtVarNm = "") As String()
ZZSql3Ly = FtLy(ZZSql3Ft(Sql3FtVarNm))
End Function

Private Function ZZSql3Ly_Wy() As WrkDr()
ZZSql3Ly_Wy = Sql3Ly_Wy(ZZSql3Ly)
End Function

Private Function ZZWy() As WrkDr()
If Sql3Ft_WrtEr(ZZSql3Ft) Then ZZSql3Ft_Edt: Stop
ZZWy = Sql3Ly_Wy(ZZSql3Ly)
End Function

Private Sub Evl__Tst(Optional Sql3FtVarNm)
If Sql3Ft_WrtEr(ZZSql3Ft(Sql3FtVarNm)) Then
    FtBrw ZZSql3Ft(Sql3FtVarNm)
    Stop
End If
Dim D As Dictionary
Set D = Evl(ZZSql3Ft(Sql3FtVarNm))
StrBrw RplVBar(D("Sql.X")), "Sql.X"
AyBrw DicLy(DicSrt(D)), "RsltDic"
Edt Sql3FtVarNm
Stop
End Sub

Private Sub Evl_1___Pass1__Tst(Optional Sql3FtVarNm = "")
L123Ay_Brw Evl_1___Pass1(ZZSql3Ly(Sql3FtVarNm))
End Sub

Sub Evl_1___Pass2__Tst(Optional Sql3FtVarNm = "")
Dim A() As L123: A = Evl_1___Pass1(ZZSql3Ly(Sql3FtVarNm))
Dim B() As L123: B = Evl_1___Pass2(A)
L123Ay_Brw B
End Sub

Function Evl_21___FmL123Ay_ToL12SOPSAy__Tst(Optional Sql3FtVarNm = "")
Dim A() As L123: A = Evl_1___Pass1(ZZSql3Ly(Sql3FtVarNm))
Dim B() As L123: B = Evl_1___Pass2(A)
Dim C() As L12SOPS: C = Evl_21___FmL123Ay_ToL12SOPSAy(B)
Dim D() As L12SOPS: D = L12SOPSAy_Srt(C)
L12SOPSAy_Brw D
End Function

Private Sub Evl_24127_Run__Tst()
Dim Dic As New Dictionary
Dim L3Prm$
Dim Ns$
    Ns = ""
    L3Prm = "CrdExpr Prm.CrdTyLis"
    Dic.Add "Prm.CrdTyLis", "1 2 3"
Dim A As KPD
    A = KPD(Ns, L3Prm, Dic)
With Evl_24127_Run(A)
    Debug.Assert .Som
    Debug.Print .Str
End With
End Sub

Private Sub Evl_241296__SqlUpd__Tst()
With Evl_241296__SqlUpd("Sql.Tx.Tx#Upd.Upd")
    Debug.Assert .Str = "Update #Tx"
    Debug.Assert .Som = True
End With
End Sub

Private Sub Evl_2412A_SqlSelFldLis__Tst()
Dim GivKey$, GivPrm$, GivDic As Dictionary
Dim ExpStr$, ExpSom As Boolean
Dim ActStr$, ActSom As Boolean
ExpSom = True
ExpStr = "    F1-Expr|    XXX 123 123|    yy           F1,|    F2-Expr      F2,|    F4           F4,|    Expr-X       X "
GivKey = "A"
GivPrm = "F1 F2 ?F3 .F4 E.X"

Set GivDic = DicByDry(Array( _
    Array("A.F1", "F1-Expr|XXX 123 123|yy"), _
    Array("A.F2", "F2-Expr"), _
    Array("A.?F3", ""), _
    Array("E.X", "Expr-X")))
Dim A As KPD
    A = KPD(GivKey, GivPrm, GivDic)
With Evl_2412A_SqlSelFldLis(A)      '<== Run
    ActSom = .Som
    ActStr = .Str
End With
Debug.Assert ActSom = ExpSom
Debug.Assert ActStr = ExpStr
End Sub

Private Sub Evl_ToDic__Tst()
Const VarNm$ = "SqlUpd"
Dim A1() As L12SOPS: A1 = ZZL12SOPSAy(VarNm)
Dim A2() As L12SOPS: A2 = Evl_SetFixStr(A1)
L12SOPSAy_Brw_NoL1L2 A2
Stop

Dim Act As Dictionary: Set Act = Evl_ToDic(A2)
DicBrw Act
End Sub

Private Sub Evl_ToDic_3_FmMulNm__Tst()
DicBrw Evl_ToDic_3_FmMulNm(ZZL12SOPSAy("SqlUpd"))
End Sub

Private Sub L12SOPSAy_Ky_UniqSrt__Tst()
AyBrw L12SOPSAy_Ky_UniqSrt(ZZL12SOPSAy)
End Sub

Private Sub Op_Sy__Tst()
AyDmp Op_Sy
End Sub

Private Sub Sql3Ft_WrtEr__Tst()
If Sql3Ft_WrtEr(ZZSql3Ft) Then ZZSql3Ft_Edt
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

Private Sub Wy_Ny__Tst()
AyBrw Wy_Ny(ZZWy)
End Sub

Private Sub ZZSql3Ly__Tst()
AyBrw ZZSql3Ly
End Sub

Sub Tst()
Tst
Op_Sy__Tst
Sql3Ly_ValidatedLy__Tst
Sql3Ft_WrtEr__Tst
Sql3Ly_Wy__Tst
ZZSql3Ly__Tst
End Sub

