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
Public AFx
Private ImpTim As Date
Const ZZOk_Fn = "Stock Card.xlsx"
Private ZZ As Permit 'for testing

Function AFxFn$()
AFxFn = FfnFn(AFx)
End Function

Sub AImport()
If IsEr Then
    Er_WrtFx
    Er_MovFx
    Log LogMsg_Er
Else
    InsTbl_Permit_and_PermitD
    Er_MovFx
    Ok_MovFx
    Log LogMsg_Ok
End If
Er_RmvPthIfEmpty
End Sub

Function Er_Ffn$()
Er_Ffn = Er_Pth & AFxFn
End Function

Sub Er_MovFx()
Fso.MoveFile AFx, Er_Ffn
End Sub

Function Er_Pth$()
Dim O$: O = PermitImpPth & "ER\": PthEns O
Er_Pth = ZMkPth(O)
End Function

Sub Er_RmvFxWs()
FxRmvWsIfExist AFx, "Error"
End Sub

Sub Er_RmvPthIfEmpty()
PthRmvEmptySubDir Er_Pth
PthRmvIfEmpty Er_Pth & "..\"
End Sub

Sub Er_WrtFx()
Dim Wb As Workbook: Set Wb = ImpPermWb
WsDlt Wb, "Error"
Dim Ws As Worksheet: Set Ws = Wb.Sheets.Add
Ws.Name = "Error"
AyPut ErMsg_Ay, WsA1(Ws)
Wb.Close True
End Sub

Function ErMsg__Bld(Sql$, QQStr$) As String()
Dim Dry: Dry = SqlDry(Sql, WrkDb)
If AyIsEmpty(Dry) Then Exit Function
Dim O$()
Dim Dr
For Each Dr In Dry
    Push O, FmtQQAv(QQStr, Dr)
Next
ErMsg__Bld = O
End Function

Function ErMsg_Ay() As String()
Dim O$()
PushAy O, ErMsg_MissingCol: If Not AyIsEmpty(O) Then ErMsg_Ay = O: Exit Function
PushAy O, ErMsg_DtaTy:  If Not AyIsEmpty(O) Then ErMsg_Ay = O: Exit Function
PushAy O, ErMsg_DupSKU
PushAy O, ErMsg_ErQty
PushAy O, ErMsg_NonIntQty
PushAy O, ErMsg_ErAmt
ErMsg_Ay = O
End Function

Function ErMsg_DtaTy() As String()
Dim O$()
PushNonEmpty O, ErMsg_DtaTy__Msg("SKU", dbText)
PushNonEmpty O, ErMsg_DtaTy__Msg("Material Description", dbText)
PushNonEmpty O, ErMsg_DtaTy__Msg("Order Qty#", dbDouble)
PushNonEmpty O, ErMsg_DtaTy__Msg("Permit Number", dbText)
PushNonEmpty O, ErMsg_DtaTy__Msg("Batch Number", dbText)
PushNonEmpty O, ErMsg_DtaTy__Msg("Duty Paid Amount", dbDouble)
ErMsg_DtaTy = O
End Function

Function ErMsg_DtaTy__Msg$(F$, ExpT As DAO.DataTypeEnum)
Dim T As DAO.DataTypeEnum: T = TblFld("Imp", F, WrkDb).Type
If T <> ExpT Then ErMsg_DtaTy__Msg = FmtQQ("Column[?] should be [?], but now it is[?]", F, DaoTyStr(ExpT), DaoTyStr(T))
End Function

Function ErMsg_DupSKU() As String()
Const S$ = "Select Distinct PermitNo,SKU,Count(*) as Cnt from Imp1 Group By PermitNo,SKU Having Count(*)>1"
Const Q$ = "There are [?] records with Same PermitNo[?] and SKU[?]"
ErMsg_DupSKU = ErMsg__Bld(S, Q)
End Function

Function ErMsg_EmptySKU() As String()
Const S$ = "Select PermitNo,Qty from Imp1 where Trim(IIf(IsNull(SKU),'',SKU))=''"
Const Q$ = "PermitNo[?] Qty[?] has Blank SKU"
ErMsg_EmptySKU = ErMsg__Bld(S, Q)
End Function

Function ErMsg_ErAmt() As String()
Const S$ = "Select PermitNo,Sku from Imp1 where Amt<=0"
Dim Dry: Dry = SqlDry(S, WrkDb)
If AyIsEmpty(Dry) Then Exit Function
Dim O$()
Dim Dr
For Each Dr In Dry
    Push O, FmtQQ("PermitNo[?] Sku[?] has Ero or negative Amount")
Next
ErMsg_ErAmt = O
End Function

Function ErMsg_ErQty() As String()
Const S$ = "Select PermitNo,Sku from Imp1 where Qty<=0"
Const Q$ = "PermitNo[?] Sku[?] has Ero or negative quantity"
ErMsg_ErQty = ErMsg__Bld(S, Q)
End Function

Function ErMsg_MissingCol() As String()
Dim Fny$(): Fny = FxWsFny(AFx, "Sheet1")
Dim Miss$()
    Dim F
    For Each F In Sy("Material Description", "Order Qty#", "Permit Number", "Batch Number", "Duty Paid Amount")
        If Not AyHas(Fny, F) Then
            Push Miss, F
        End If
    Next
If AyIsEmpty(Miss) Then Exit Function
Dim O$()
    Push O, "Missing column: " & JnComma(AyQuote(Miss, "[]"))
ErMsg_MissingCol = O
End Function

Function ErMsg_NonIntQty() As String()
Const S$ = "Select PermitNo,Sku,Qty from Imp1 where Int(Qty)<>Qty"
Const Q$ = "PermitNo[?] Sku[?] has non-integer quantity[?]"
ErMsg_NonIntQty = ErMsg__Bld(S, Q)
End Function

Function ErMsg_Sheet1() As String()
'Should only 1 sheet: Sheet1
'Or 2 sheets        : Sheet1 and Error
Dim WsNy$(): WsNy = FxWsNy(AFx)
Dim NWs%: NWs = Sz(WsNy)
Dim O$()
    If NWs <> 1 And NWs <> 2 Then
        Push O, FmtQQ("The workboOk_ should have 1 or 2 worksheets, but now it is [?] worksheets: " & JnSpc(AyQuote(WsNy, "[]")), NWs)
        ErMsg_Sheet1 = O
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
ErMsg_Sheet1 = O
End Function

Function ImpPermDrs() As Drs
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
    WsClsNoSav Ws
    X = SqSel(Sq, MapStr)
End If
ImpPermDrs = X
End Function

Function ImpPermHdrDrs() As Drs
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

Function ImpPermItmDrs() As Drs
'Oup: Drs: Permit SKU SeqNo Qty BchNo Amt
'Inp: Imp1: SKU Qty PermitNo BchNo Amt
Dim Dry(): Dry = SqlDry("Select SKU,Qty,PermitNo,BchNo,Amt from Imp1 Order By PermitNo", WrkDb)
Dim ODry()
    Dim ODr()
    Dim Sku$, Qty@, PermitNo$, BchNo$, Amt#, SeqNo%
    Dim Dr, LasPermitNo
    For Each Dr In Dry
        SeqNo = SeqNo + 10
        AyAsg Dr, Sku, Qty, PermitNo, BchNo, Amt
        If LasPermitNo <> PermitNo Then
            LasPermitNo = PermitNo
            SeqNo = 10
        End If
        Erase ODr
        Push ODr, ZDicOfId_ByPermitNo(PermitNo)
        Push ODr, Sku
        Push ODr, SeqNo
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

Function ImpPermWb() As Workbook
Set ImpPermWb = FxOpn(AFx)
End Function

Function ImpPermWs() As Worksheet
Set ImpPermWs = ImpPermWb.Sheets(1)
End Function

Sub InsTbl_Permit_and_PermitD()
RunSqlAy SqlAy_UpdTbl_Permit_and_PermitD, WrkDb
End Sub

Function IsEr() As Boolean
IsEr = Not AyIsEmpty(ErMsg_Ay)
End Function

Function LogMsg_Er$()
End Function

Function LogMsg_Ok$()
LogMsg_Ok = "sdff"
End Function

Function Ok_Ffn$()
Ok_Ffn = Ok_Pth & AFxFn
End Function

Sub Ok_MovFx()
Fso.MoveFile AFx, Ok_Ffn
End Sub

Function Ok_Pth$()
Dim O$: O = PermitImpPth & "Ok_\": PthEns O
Ok_Pth = ZMkPth(O)
End Function

Function PermitNoAy() As String()
Static Z As Boolean, X
If Not Z Then
    Z = True
    X = DrsStrCol(ImpPermHdrDrs, "PermitNo")
End If
PermitNoAy = X
End Function

Function PermitNoLis$()
PermitNoLis = JnComma(AyQuote(PermitNoAy, "'"))
End Function

Function Sql_DltTbl_Permit$()
Const S$ = "Delete From Permit Where PermitNo in (?)"
Sql_DltTbl_Permit = FmtQQ(S, PermitNoLis)
End Function

Function Sql_DltTbl_PermitD$()
Const S$ = "Delete From PermitD Where Permit in (Select Permit from Permit where PermitNo in (?))"
Sql_DltTbl_PermitD = FmtQQ(S, PermitNoLis)
End Function

Function SqlAy_InsTbl_Permit() As String()
Dim ODry()
    Dim M As PermitDftVal: M = PermitDftVal
    Dim Dr
    Dim ODr()
    Dim Dry()
    Dim Drs As Drs: Drs = ImpPermHdrDrs
    Dim IdxAy&(): IdxAy = FidxAy(Drs.Fny, "PermitNo NSKU Qty Amt")
    Dim PermitNo$, NSku%, Qty&, Amt#
    For Each Dr In Drs.Dry
        AyAsg_Idx Dr, IdxAy, PermitNo, NSku, Qty, Amt
        Erase ODr
        Push ODr, ZDicOfId_ByPermitNo(Dr(0))
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
SqlAy_InsTbl_Permit = O
End Function

Function SqlAy_InsTbl_PermitD() As String()
'PermitD = * Permit Sku SeqNo Qty BchNo Rate Amt DteCrt
Dim ODry()
    Dim M As PermitDftVal: M = PermitDftVal
    Dim Dr
    Dim ODr()
    Dim Dry()
    Dim Id&: Id = TblNxtId("PermitD")
    Dim Permit&, Sku$, SeqNo%, Qty&, BchNo$, Amt@
    Dim Drs As Drs: Drs = ImpPermItmDrs
    Dim IdxAy&(): IdxAy = FidxAy(Drs.Fny, "Permit SKU SeqNo Qty BchNo Amt")
    For Each Dr In ImpPermItmDrs.Dry
        AyAsg_Idx Dr, IdxAy, Permit, Sku, SeqNo, Qty, BchNo, Amt
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
SqlAy_InsTbl_PermitD = O
End Function

Function SqlAy_UpdTbl_Permit_and_PermitD() As String()
Dim O$()
Push O, Sql_DltTbl_Permit
Push O, Sql_DltTbl_PermitD
PushAy O, SqlAy_InsTbl_Permit
PushAy O, SqlAy_InsTbl_PermitD
SqlAy_UpdTbl_Permit_and_PermitD = O
End Function

Function WrkDb() As Database
Static O As Database
If IsNothing(O) Then
    Set O = TmpDb
    LnkFx AFx, "Sheet1", "Imp", O
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

Function ZDicOfId_ByPermitNo() As Dictionary
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
Set ZDicOfId_ByPermitNo = X
End Function

Function ZMkPth$(P)
Dim O$: O = P & Format(ImpTim, "YYYYMMDD_HHMMSS") & "\"
PthEns O
ZMkPth = O
End Function

Function ZZResPth$()
ZZResPth = PjSrcPth & "TstRes\"
End Function

Sub ZZSet(FxFn)
Set ZZ = New Permit
ZZ.AFx = ZZResPth & FxFn
End Sub

Private Sub Class_Initialize()
ImpTim = Now
End Sub

Private Sub Class_Terminate()
On Error GoTo X
Dim Fb$: Fb = WrkDb.Name
WrkDb.Close
Kill Fb
Log FmtQQ("TmpFb[?] killed", Fb)
X:
End Sub

Sub ErMsg_Ay__Tst()
ZZSet ZZOk_Fn
AyDmp ZZ.ErMsg_Ay
End Sub

Sub ErMsg_DtaTy__Tst()
Dim Exp1$(), Exp2$(), Exp3$()
Push Exp1, "Column[SKU] should be [Text], but now it is[Double]"
ZZSet "ErDtaTy_Fx1.xls": AssertEqAy Exp1, ZZ.ErMsg_DtaTy
End Sub

Sub ErMsg_DupSKU__Tst()
Dim Exp1$()
ZZSet "ErDupSKU.xls": AssertEqAy Exp1, ZZ.ErMsg_DupSKU
End Sub

Sub ErMsg_EmptySKU__Tst()
Dim Exp$()
Push Exp, "PermitNo[Permit1] Qty[1] has Blank SKU"
ZZSet "ErEmptySKU.xls": AssertEqAy Exp, ZZ.ErMsg_EmptySKU
End Sub

Sub ErMsg_MissingCol__Tst()
Dim Exp$()
Push Exp, "Missing column: [Material Description],[Order Qty#],[Permit Number],[Batch Number],[Duty Paid Amount]"
ZZSet "ErMissingCol.xls": AssertEqAy Exp, ZZ.ErMsg_MissingCol
End Sub

Sub ErMsg_NonIntQty__Tst()
Dim Exp$()
Push Exp, "PermitNo[Permit1] Sku[Sku-1] has non-integer quantity[1.1]"
ZZSet "ErNonIntQty.xls": AssertEqAy Exp, ZZ.ErMsg_NonIntQty
End Sub

Sub ErMsg_Sheet1__Tst()
Dim Exp1$(), Exp2$(), Exp3$()
Push Exp1, "The workboOk_ should have 1 or 2 worksheets, but now it is [3] worksheets: [Sheet1] [Sheet2] [Sheet3]"
Push Exp2, "Worksheet[Sheet1] not found"
Push Exp3, "The 2 worksheet name should be [Sheet1] and [Error], but now they are [Sheet1] and [Sheet2]"
ZZSet "ErSheet1_Fx1.xls": AssertEqAy Exp1, ZZ.ErMsg_Sheet1
ZZSet "ErSheet1_Fx2.xls": AssertEqAy Exp2, ZZ.ErMsg_Sheet1
ZZSet "ErSheet1_Fx3.xls": AssertEqAy Exp3, ZZ.ErMsg_Sheet1
End Sub

Sub ImpPermItmDrs__Tst()
ZZSet ZZOk_Fn
DrsBrw ZZ.ImpPermItmDrs
End Sub

Function SqlAy_InsTbl_Permit__Tst()
ZZSet ZZOk_Fn
AyBrw ZZ.SqlAy_InsTbl_Permit
End Function

Sub SqlAy_InsTbl_PermitD__Tst()
ZZSet ZZOk_Fn
AyBrw ZZ.SqlAy_InsTbl_PermitD
End Sub

Sub SqlAy_UpdTbl_Permit_and_PermitD__Tst()
ZZSet ZZOk_Fn
AyBrw ZZ.SqlAy_UpdTbl_Permit_and_PermitD
End Sub

Sub WrkDb__Tst()
ZZSet ZZOk_Fn
Dim D As Database: Set D = ZZ.WrkDb
Dim Fb$: Fb = D.Name
Dim mTny$(): mTny = DbTny(D)
Debug.Assert Sz(mTny) = 4
AssertEqAy SplitSpc("Permit PermitD Imp Imp1"), mTny
End Sub

Sub Tst()
ErMsg_Ay__Tst
ErMsg_DtaTy__Tst
ErMsg_DupSKU__Tst
ErMsg_EmptySKU__Tst
ErMsg_MissingCol__Tst
ErMsg_NonIntQty__Tst
ErMsg_Sheet1__Tst
ImpPermItmDrs__Tst
SqlAy_InsTbl_Permit__Tst
SqlAy_InsTbl_PermitD__Tst
SqlAy_UpdTbl_Permit_and_PermitD__Tst
WrkDb__Tst
End Sub

