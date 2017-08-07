VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Permit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private Fx
Private ImpTim As Date
Private Const Tst_OkFn = "Stock Card.xlsx"
Private Const Tst_ErFn = "Error.xlsx"

Private Function Tst_OkFx$()
Tst_OkFx = Tst_ResPth & Tst_OkFn
End Function

Private Function Tst_ErFfn$()
Tst_ErFfn = Tst_ResPth & Tst_ErFn
End Function

Private Sub Class_Initialize()
ImpTim = Now
End Sub

Sub TstOk()
Fx = Tst_OkFx
WrkDb__Tst
End Sub
Sub TstEr()
'Er_MissingCol__Tst
Er_EmptySKU__Tst
'Er_NonIntQty__Tst
'Er_DtaTy__Tst
'Er_Sheet1__Tst
End Sub
Private Function Tst_FxAy() As String()
Tst_FxAy = PthFfnAy(PermitImpPth)
End Function
Private Sub Tst_RstOkFx()
FilCpyToPth Tst_OkFx, PermitImpPth
End Sub
Private Function Tst_ResPth$()
Tst_ResPth = PjSrcPth & "TstRes\"
End Function

Private Sub WrkDb__Tst()
Dim D As Database: Set D = WrkDb
Dim Fb$: Fb = D.Name
Dim mTny$(): mTny = Tny(D)
Debug.Assert Sz(mTny) = 4
AssertEqAy SplitSpc("Permit PermitD Imp Imp1"), mTny
End Sub

Sub Impport(P_Fx)
Fx = P_Fx
If IsEr Then
    WrtER
    MovToER
    Log MsgEr
Else
    Dim A As Boolean
    A = CallByName(Me, "IsEr", VbMethod)
    InsTbl_Permit_and_PermitD
    RmvWsEr
    MovToOK
    Log MsgOK
End If
RmvErPthIfEmpty
End Sub
Private Sub InsTbl_Permit_and_PermitD()
RunSqlAy InsTbl_Permit_and_PermitD_SqlAy, WrkDb
End Sub
Private Sub RstTstFfn()
If HasFil(PermitImpPth) Then Exit Sub
FileCopy TstFfn, PermitImpPth & FfnFn(TstFfn)
End Sub

Private Function TstFfn$()
TstFfn = PermitImpPth & "Tst\Stock card.xlsx"
End Function

Private Sub RmvErPthIfEmpty()
RmvPthIfEmpty ErPth & "..\"
End Sub

Private Function HasErWs() As Boolean
Dim WsNy$(): WsNy = FxWsNy(Fx)
HasErWs = AyHas(WsNy, "Error")
End Function

Private Sub RmvWsEr()
If Not HasErWs Then Exit Sub
Dim Wb As Workbook: Set Wb = ImpPermWb
DltWs Wb, "Error"
Wb.Close True
End Sub

Private Sub ErMsgAy__Tst()
AyDmp ErMsgAy
End Sub

Private Function ErMsgAy() As String()
Dim O$()
PushAy O, Er_MissingCol: If Not IsEmptyAy(O) Then ErMsgAy = O: Exit Function
PushAy O, Er_DtaTy:  If Not IsEmptyAy(O) Then ErMsgAy = O: Exit Function
PushAy O, Er_DupSKU
PushAy O, Er_ZerQty
PushAy O, Er_NonIntQty
PushAy O, Er_ZerAmt
ErMsgAy = O
End Function

Private Function Er_ZerAmt() As String()
Const S$ = "Select PermitNo,Sku from Imp1 where Amt<=0"
Dim Dry: Dry = SqlDry(S, WrkDb)
If IsEmptyAy(Dry) Then Exit Function
Dim O$()
Dim Dr
For Each Dr In Dry
    Push O, FmtQQ("PermitNo[?] Sku[?] has zero or negative Amount")
Next
Er_ZerAmt = O
End Function

Private Function WrkDb() As Database
Static O As Database
If IsNothing(O) Then
    Set O = TmpDb
    LnkFx Fx, "Sheet1", "Imp", O
    LnkFb DtaFb, "Permit", , O
    LnkFb DtaFb, "PermitD", , O
    RunSql "Select " & _
        "SKU," & _
        "[Material Description] As Des," & _
        "[Order Qty#]           As Qty," & _
        "[Permit Number]        As PermitNo," & _
        "[Batch Number]         As BchNo," & _
        "[Duty Paid Amount]     As Amt" & _
        " Into Imp1 from Imp", O
End If
Set WrkDb = O
End Function


Private Function Er_MissingCol() As String()
Dim Fny$(): Fny = FxWsFny(Fx, "Sheet1")
Dim Miss$()
    Dim F
    For Each F In Sy("Material Description", "Order Qty#", "Permit Number", "Batch Number", "Duty Paid Amount")
        If Not AyHas(Fny, F) Then
            Push Miss, F
        End If
    Next
If IsEmptyAy(Miss) Then Exit Function
Dim O$()
    Push O, "Missing column: " & JnComma(QuoteAy(Miss, "[]"))
Er_MissingCol = O
End Function

Private Function Er_DtaTyMsg$(F$, ExpT As DAO.DataTypeEnum)
Dim T As DAO.DataTypeEnum: T = TblFld("Imp", F, WrkDb).Type
If T <> ExpT Then Er_DtaTyMsg = FmtQQ("Column[?] should be [?], but now it is[?]", F, DaoTyStr(ExpT), DaoTyStr(T))
End Function

Private Sub Er_DtaTy__Tst()
Dim Exp1$(), Exp2$(), Exp3$()
Push Exp1, "Column[SKU] should be [Text], but now it is[Double]"
Fx = Tst_ResPth & "ErDtaTy_Fx1.xls": AssertEqAy Exp1, Er_DtaTy
End Sub

Private Sub Er_MissingCol__Tst()
Dim Exp$()
Push Exp, "Missing column: [Material Description],[Order Qty#],[Permit Number],[Batch Number],[Duty Paid Amount]"
Fx = Tst_ResPth & "ErMissingCol.xls": AssertEqAy Exp, Er_MissingCol
End Sub

Private Sub Er_Sheet1__Tst()
Dim Exp1$(), Exp2$(), Exp3$()
Push Exp1, "The workbook should have 1 or 2 worksheets, but now it is [3] worksheets: [Sheet1] [Sheet2] [Sheet3]"
Push Exp2, "Worksheet[Sheet1] not found"
Push Exp3, "The 2 worksheet name should be [Sheet1] and [Error], but now they are [Sheet1] and [Sheet2]"
Fx = Tst_ResPth & "ErSheet1_Fx1.xls": AssertEqAy Exp1, Er_Sheet1
Fx = Tst_ResPth & "ErSheet1_Fx2.xls": AssertEqAy Exp2, Er_Sheet1
Fx = Tst_ResPth & "ErSheet1_Fx3.xls": AssertEqAy Exp3, Er_Sheet1
End Sub

Private Function Er_Sheet1() As String()
'Should only 1 sheet: Sheet1
'Or 2 sheets        : Sheet1 and Error
Dim WsNy$(): WsNy = FxWsNy(Fx)
Dim NWs%: NWs = Sz(WsNy)
Dim O$()
    If NWs <> 1 And NWs <> 2 Then
        Push O, FmtQQ("The workbook should have 1 or 2 worksheets, but now it is [?] worksheets: " & JnSpc(QuoteAy(WsNy, "[]")), NWs)
        Er_Sheet1 = O
        Exit Function
    End If
    If NWs = 1 Then
        If WsNy(0) <> "Sheet1" Then
            Push O, "Worksheet[Sheet1] not found"
        End If
    End If
    If NWs = 2 Then
        If (WsNy(0) <> "Error" And WsNy(0) <> "Sheet1") Or (WsNy(1) <> "Error" And WsNy(1) <> "Sheet1") Then
            Push O, FmtQQ("The 2 worksheet name should be [Sheet1] and [Error], but now they are [?] and [?]", WsNy(0), WsNy(1))
        End If
    End If
Er_Sheet1 = O
End Function

Private Function Er_DtaTy() As String()
Dim O$()
PushNonEmpty O, Er_DtaTyMsg("SKU", dbText)
PushNonEmpty O, Er_DtaTyMsg("Material Description", dbText)
PushNonEmpty O, Er_DtaTyMsg("Order Qty#", dbDouble)
PushNonEmpty O, Er_DtaTyMsg("Permit Number", dbText)
PushNonEmpty O, Er_DtaTyMsg("Batch Number", dbText)
PushNonEmpty O, Er_DtaTyMsg("Duty Paid Amount", dbDouble)
Er_DtaTy = O
End Function

Private Sub Er_DupSKU__Tst()
AyDmp Er_DupSKU()
End Sub

Private Function Er_DupSKU() As String()
Const S$ = "Select Distinct PermitNo,SKU,Count(*) as Cnt from Imp1 Group By PermitNo,SKU Having Count(*)>1"
Const Q$ = "There are [?] records with Same PermitNo[?] and SKU[?]"
Er_DupSKU = Er__MsgAy(S, Q)
End Function

Private Function Er_ZerQty() As String()
Const S$ = "Select PermitNo,Sku from Imp1 where Qty<=0"
Const Q$ = "PermitNo[?] Sku[?] has zero or negative quantity"
Er_ZerQty = Er__MsgAy(S, Q)
End Function

Private Function Er_NonIntQty() As String()
Const S$ = "Select PermitNo,Sku,Qty from Imp1 where Int(Qty)<>Qty"
Const Q$ = "PermitNo[?] Sku[?] has non-integer quantity[?]"
Er_NonIntQty = Er__MsgAy(S, Q)
End Function

Private Function Er__MsgAy(Sql$, QQStr$) As String()
Dim Dry: Dry = SqlDry(Sql, WrkDb)
If IsEmptyAy(Dry) Then Exit Function
Dim O$()
Dim Dr
For Each Dr In Dry
    Push O, FmtQQAv(QQStr, Dr)
Next
Er__MsgAy = O
End Function

Private Function Er_EmptySKU() As String()
Const S$ = "Select PermitNo,Qty from Imp1 where Trim(IIf(IsNull(SKU),'',SKU))=''"
Const Q$ = "PermitNo[?] Qty[?] has Blank SKU"
Er_EmptySKU = Er__MsgAy(S, Q)
End Function

Private Sub Er_EmptySKU__Tst()
Dim Exp$()
Push Exp, "PermitNo[Permit1] Qty[1] has Blank SKU"
Fx = Tst_ResPth & "ErEmptySKU.xls": AssertEqAy Exp, Er_EmptySKU
End Sub

Private Sub Er_NonIntQty__Tst()
Dim Exp$()
Push Exp, "PermitNo[Permit1] Sku[Sku-1] has non-integer quantity[1.1]"
Fx = Tst_ResPth & "ErNonIntQty.xls": AssertEqAy Exp, Er_NonIntQty
End Sub

Function IsEr() As Boolean
IsEr = Not IsEmptyAy(ErMsgAy)
End Function

Private Sub MovToER()
Fso.MoveFile Fx, ErFfn
End Sub

Private Sub MovToOK()
Fso.MoveFile Fx, OkFfn
End Sub

Private Function OkPth$()
Dim O$: O = PermitImpPth & "OK\": PthEns O
OkPth = MkPth(O)
End Function

Private Function MkPth$(P)
Dim O$: O = P & Format(ImpTim, "YYYYMMDD_HHMMSS") & "\"
PthEns O
MkPth = O
End Function

Private Function ErPth$()
Dim O$: O = PermitImpPth & "ER\": PthEns O
ErPth = MkPth(O)
End Function

Private Function FxFn$()
FxFn = FfnFn(Fx)
End Function

Private Function OkFfn$()
OkFfn = OkPth & FxFn
End Function

Private Function ErFfn$()
ErFfn = ErPth & FxFn
End Function

Private Sub WrtER()
Dim Wb As Workbook: Set Wb = ImpPermWb
DltWs Wb, "Error"
Dim Ws As Worksheet: Set Ws = Wb.Sheets.Add
Ws.Name = "Error"
AyPut ErMsgAy, WsA1(Ws)
Wb.Close True
End Sub
Private Sub ImpPermItmDrs__Tst()
Dim Exp$()
Dim Act$()
Fx = Tst_ResPth & "Stock Card.xls": AssertEqAy Exp, Act
End Sub
Private Function ImpPermItmDrs() As Drs
'Oup: Drs: Permit SKU SeqNo Qty BchNo Amt
'Inp: Imp1: SKU Qty PermitNo BchNo Amt
Dim Dry(): Dry = SqlDry("Select SKU,Qty,PermitNo,BchNo,Amt from Imp1", WrkDb)
Dim ODry()
    Dim ODr()
    Dim Sku$, Qty@, PermitNo$, BchNo$, Amt#, SeqNo%
    Dim Dr, LasPermitNo
    For Each Dr In Dry
        AyAsg Dr, Sku, Qty, PermitNo, BchNo, Amt
        If LasPermitNo <> PermitNo Then
            LasPermitNo = PermitNo
            SeqNo = 0
        End If
        Erase ODr
        Push ODr, DicOfId_ByPermitNo(PermitNo)
        Push ODr, Sku
        Push ODr, SeqNo: SeqNo = SeqNo + 1
        Push ODr, Qty
        Push ODr, BchNo
        Push ODr, Amt
        Push ODry, ODr
    Next
Dim O As Drs
    O.Fny = SplitSpc("Permit SKU SeqNo Qty BchNo Amt")
    O.Dry = ODry
ImpPermItmDrs = O
'Inp:
'|---------|-----------------------------------------|-------|----------------|------------|-----|
'| SKU     | Des                                     | Qty   | PermitNo       | BchNo      | Amt |
'|---------|-----------------------------------------|-------|----------------|------------|-----|
'| 1056951 | HY-VSOPP40-F0,2L--C24-HNG DP-M12        | 1680  | E2016006258RLA | 2016006258 | $   |
End Function

Private Function ImpPermHdrDrs() As Drs
'Summing following in WrkDb->Imp1
Const Sql$ = "Select Distinct PermitNo,Count(*) as NSKU,Sum(x.Qty) as Qty,Sum(x.Amt) as Amt from Imp1 x Group By PermitNo"
ImpPermHdrDrs = SqlDrs(Sql, WrkDb)
'PermitNo NSKU Qty Amt
'|---------|-----------------------------------------|-------|----------------|------------|-----|
'| SKU     | Des                                     | Qty   | PermitNo       | BchNo      | Amt |
'|---------|-----------------------------------------|-------|----------------|------------|-----|
'| 1056951 | HY-VSOPP40-F0,2L--C24-HNG DP-M12        | 1680  | E2016006258RLA | 2016006258 | $   |
'| 1056951 | HY-VSOPP40-F0,2L--C24-HNG DP-M12        | 2016  | E2016009955RLA | 2016009955 | $   |
'| 1056951 | HY-VSOPP40-F0,2L--C24-HNG DP-M12        | 1104  | E2016009959RLA | 2016009959 | $   |
'| 1056819 | HY-VSOPP 40-0,35L-WB-C12-HNG DP M12     | 1188  | E2016021484RLA | 2016021484 | $   |
'| 1056819 | HY-VSOPP 40-0,35L-WB-C12-HNG DP M12     | 1188  | E2016021491RLA | 2016021491 | $   |
'| 1055625 | HY-VSOPP40-0,7L-WB-C12-HNG DP-M12       | 1404  | E2016021488RLA | 2016021488 | $   |
'| 1055625 | HY-VSOPP40-0,7L-WB-C12-HNG DP-M12       | 16128 | E2016021489RLA | 2016021489 | $   |
'| 1055642 | HY-VSOPP40-0,7L--C12-HNG DP M12         | 2688  | E2016021485RLA | 2016021485 | $   |
'| 1055642 | HY-VSOPP40-0,7L--C12-HNG DP M12         | 11712 | E2016021487RLA | 2016021487 | $   |
'| 1055651 | HY-VSOPP40-1,5L-WB-C6-HNG DP-M12        | 60    | E2016018313RLA | 2016018313 | $   |
'| 1055651 | HY-VSOPP40-1,5L-WB-C6-HNG DP-M12        | 360   | E2016019608RLA | 2016019608 | $   |
'| 1055651 | HY-VSOPP40-1,5L-WB-C6-HNG DP-M12        | 60    | E2016019610RLA | 2016019610 | $   |
'| 1055651 | HY-VSOPP40-1,5L-WB-C6-HNG DP-M12        | 120   | E2016020569RLA | 2016020569 | $   |
'| 1073359 | HY-VSOPP40-0,70-LGB2G-C6-HNG DP-LED2016 | 60    | E2016021491RLA | 2016021491 | $   |
'| 1054852 | HY-XO40-0,7L-WB-C12-HNG DP-M12          | 1200  | E2016020571RLA | 2016020571 | $   |
'| 1054852 | HY-XO40-0,7L-WB-C12-HNG DP-M12          | 3600  | E2016021484RLA | 2016021484 | $   |
'| 1055261 | HY-XO40-1,5L-WB-C6-HNG DP-M12           | 60    | E2016019608RLA | 2016019608 | $   |
'| 1055261 | HY-XO40-1,5L-WB-C6-HNG DP-M12           | 216   | E2016019610RLA | 2016019610 | $   |
'| 1055261 | HY-XO40-1,5L-WB-C6-HNG DP-M12           | 204   | E2016020565RLA | 2016020565 | $   |
'| 1061270 | HY-PARAD40-0,7L-LGB-C3-HNG DP M13       | 180   | E2016021483RLA | 2016021483 |     |
'|---------|-----------------------------------------|-------|----------------|------------|-----|
End Function

Private Function Ins_TblPermitD_SqlAy() As String()
'PermitD = * Permit Sku SeqNo Qty BchNo Rate Amt DteCrt
Dim ODry()
    Dim M As PermitDftVal: M = PermitDftVal
    Dim Dr
    Dim ODr()
    Dim Dry()
    Dim Id&: Id = TblNxtId("PermitD")
    Dim Permit&, Sku$, SeqNo%, Qty&, BchNo$, Amt@
    For Each Dr In ImpPermItmDrs.Dry 'Permit SKU SeqNo Qty BchNo Amt
                                     '0      1   2     3   4     5
        AyAsg Dr, Permit, Sku, SeqNo, Qty, BchNo, Amt
        Erase ODr
        Push ODr, Id: Id = Id + 1
        Push ODr, Permit
        Push ODr, Sku
        Push ODr, SeqNo
        Push ODr, Qty
        Push ODr, Amt / Qty
        Push ODr, BchNo
        Push ODr, Amt
        Push ODr, Now
        Push ODry, ODr
    Next
Dim Dt As Dt
    Dt.Fny = SplitSpc("PermitD Permit Sku SeqNo Qty BchNo Rate Amt DteCrt")
    Dt.Dry = ODry
    Dt.DtNm = "PermitD"
Dim O$()
    O = InsDtSqlAy(Dt)
AyBrw O
Ins_TblPermitD_SqlAy = O
End Function

Private Function Ins_TblPermit_SqlAy() As String()
Dim ODry()
    Dim M As PermitDftVal: M = PermitDftVal
    Dim Dr
    Dim ODr()
    Dim Dry()
    Dim PermitNo$, NSku%, Qty&, Amt#
    For Each Dr In ImpPermHdrDrs.Dry 'PermitNo NSKU Qty Amt
        AyAsg Dr, PermitNo, NSku, Qty, Amt
        Erase ODr
        Push ODr, DicOfId_ByPermitNo(Dr(0))
        Push ODr, PermitNo ' PermitNo
        Push ODr, Format(ImpTim, "YYYY-MM-DD")
        Push ODr, Format(ImpTim, "YYYY-MM-DD")
        Push ODr, NSku
        Push ODr, Qty
        Push ODr, Amt
        Push ODr, M.GLAc
        Push ODr, M.GLAcName
        Push ODr, M.BankCode
        Push ODr, M.ByUsr
        Push ODr, Now
        Push ODr, True
        Push ODry, ODr
    Next
Dim Dt As Dt
    Dt.Fny = SplitSpc("Permit PermitNo PermitDate PostDate NSku Qty Tot GLAc GLAcName BankCode ByUsr DteCrt IsImport")
    Dt.Dry = ODry
    Dt.DtNm = "Permit"
Dim O$()
    O = InsDtSqlAy(Dt)
Ins_TblPermit_SqlAy = O
End Function

Private Function InsTbl_Permit_and_PermitD_SqlAy() As String()
Dim O$()
Push O, Dlt_TblPermit_Sql
Push O, Dlt_TblPermitD_Sql
PushAy O, Ins_TblPermit_SqlAy
PushAy O, Ins_TblPermitD_SqlAy
InsTbl_Permit_and_PermitD_SqlAy = O
End Function

Private Function DicOfId_ByPermitNo() As Dictionary
Static Z As Boolean, X As Dictionary
If Not Z Then
    Z = True
    Set X = New Dictionary
    Dim Id&: Id = TblNxtId("Permit")
    Dim I
    For Each I In PermitNoAy
        X.Add I, Id: Id = Id + 1
    Next
End If
Set DicOfId_ByPermitNo = X
End Function

Private Function Dlt_TblPermit_Sql$()
Const S$ = "Delete From Permit Where PermitNo in (?)"
Dlt_TblPermit_Sql = FmtQQ(S, PermitNoLis)
End Function
Private Function Dlt_TblPermitD_Sql$()
Const S$ = "Delete From PermitD Where Permit in (Select Permit from Permit where PermitNo in (?))"
Dlt_TblPermitD_Sql = FmtQQ(S, PermitNoLis)
End Function
Private Function PermitNoLis$()
PermitNoLis = JnComma(QuoteAy(PermitNoAy, "'"))
End Function
Private Function PermitNoAy() As String()
Static Z As Boolean, X
If Not Z Then
    Z = True
    X = DrsStrCol(ImpPermHdrDrs, "PermitNo")
End If
PermitNoAy = X
End Function
Private Function TblPemitD_SqlAy() As String()
'*Tbl PermitD
'|---------|--------|----------|-------|-------|------------|------------|--------------|------------------------|------------------------|
'| PermitD | Permit | Sku      | SeqNo | Qty   | BchNo      | Rate       | Amt          | DteCrt                 | DteUpd                 |
'|---------|--------|----------|-------|-------|------------|------------|--------------|------------------------|------------------------|

'PermitD = * Permit Sku SeqNo Qty BchNo Rate Amt DteCrt DteUpd
Dim O As Dt
    O.Fny = SplitSpc("PermitD ")
If Not IsEr Then
    Dim M As TDrImpPermItm
    Dim Seq&: Seq = 0
    Dim I
'    For Each I In ImpPermDry
'        Set M = I
'        Dim Dr As New DrPermitD
'            Dim Dte As Date: Dte = Date
'            With Dr
'                Seq = Seq + 10
'                .Permit = The_PermitId
'                .Sku = M.Sku
'                .SeqNo = Seq
'                .Qty = M.Qty
'                .BchNo = M.BchNo
'                .Rate = M.Rate
'                .Amt = M.Amt
'                .DteCrt = Now
'                '.DteUpd As Date
'            End With
'        AddDr O, Dr
'    Next
End If
'TblPemitD_SqlAy = O
End Function
Private Function ImpPermWs() As Worksheet
Set ImpPermWs = ImpPermWb.Sheets(1)
End Function
Private Function ImpPermWb() As Workbook
Set ImpPermWb = OpnWb(Fx)
End Function

Private Function ImpPermDrs() As Drs
'|---------|-----------------------------------------|-------|----------------|------------|-----|
'| SKU     | Des                                     | Qty   | PermitNo       | BchNo      | Amt |
'|---------|-----------------------------------------|-------|----------------|------------|-----|
'| 1056951 | HY-VSOPP40-F0,2L--C24-HNG DP-M12        | 1680  | E2016006258RLA | 2016006258 | $   |
'| 1056951 | HY-VSOPP40-F0,2L--C24-HNG DP-M12        | 2016  | E2016009955RLA | 2016009955 | $   |
'| 1056951 | HY-VSOPP40-F0,2L--C24-HNG DP-M12        | 1104  | E2016009959RLA | 2016009959 | $   |
'| 1056819 | HY-VSOPP 40-0,35L-WB-C12-HNG DP M12     | 1188  | E2016021484RLA | 2016021484 | $   |
'| 1056819 | HY-VSOPP 40-0,35L-WB-C12-HNG DP M12     | 1188  | E2016021491RLA | 2016021491 | $   |
'| 1055625 | HY-VSOPP40-0,7L-WB-C12-HNG DP-M12       | 1404  | E2016021488RLA | 2016021488 | $   |
'| 1055625 | HY-VSOPP40-0,7L-WB-C12-HNG DP-M12       | 16128 | E2016021489RLA | 2016021489 | $   |
'| 1055642 | HY-VSOPP40-0,7L--C12-HNG DP M12         | 2688  | E2016021485RLA | 2016021485 | $   |
'| 1055642 | HY-VSOPP40-0,7L--C12-HNG DP M12         | 11712 | E2016021487RLA | 2016021487 | $   |
'| 1055651 | HY-VSOPP40-1,5L-WB-C6-HNG DP-M12        | 60    | E2016018313RLA | 2016018313 | $   |
'| 1055651 | HY-VSOPP40-1,5L-WB-C6-HNG DP-M12        | 360   | E2016019608RLA | 2016019608 | $   |
'| 1055651 | HY-VSOPP40-1,5L-WB-C6-HNG DP-M12        | 60    | E2016019610RLA | 2016019610 | $   |
'| 1055651 | HY-VSOPP40-1,5L-WB-C6-HNG DP-M12        | 120   | E2016020569RLA | 2016020569 | $   |
'| 1073359 | HY-VSOPP40-0,70-LGB2G-C6-HNG DP-LED2016 | 60    | E2016021491RLA | 2016021491 | $   |
'| 1054852 | HY-XO40-0,7L-WB-C12-HNG DP-M12          | 1200  | E2016020571RLA | 2016020571 | $   |
'| 1054852 | HY-XO40-0,7L-WB-C12-HNG DP-M12          | 3600  | E2016021484RLA | 2016021484 | $   |
'| 1055261 | HY-XO40-1,5L-WB-C6-HNG DP-M12           | 60    | E2016019608RLA | 2016019608 | $   |
'| 1055261 | HY-XO40-1,5L-WB-C6-HNG DP-M12           | 216   | E2016019610RLA | 2016019610 | $   |
'| 1055261 | HY-XO40-1,5L-WB-C6-HNG DP-M12           | 204   | E2016020565RLA | 2016020565 | $   |
'| 1061270 | HY-PARAD40-0,7L-LGB-C3-HNG DP M13       | 180   | E2016021483RLA | 2016021483 |     |
'|---------|-----------------------------------------|-------|----------------|------------|-----|
Const MapStr$ = _
"SKU" & _
"|Des     :Material Description" & _
"|Qty     :Order Qty." & _
"|PermitNo:Permit Number" & _
"|BchNo   :Batch Number" & _
"|Amt     :Duty Paid Amount"
Dim X As Drs, Z As Boolean
If Not Z Then
    Dim Ws As Worksheet: Set Ws = ImpPermWs
    Z = True
    Dim Sq: Sq = WsSq(ImpPermWs)
    ClsWsNoSav Ws
    X = SqSel(Sq, MapStr)
End If
ImpPermDrs = X
End Function

Private Function MsgOK$()
MsgOK = "sdff"
End Function

Private Function MsgEr$()
End Function

Private Sub Class_Terminate()
On Error GoTo X
Dim Fb$: Fb = WrkDb.Name
WrkDb.Close
Kill Fb
Log FmtQQ("TmpFb[?] killed", Fb)
X:
End Sub
