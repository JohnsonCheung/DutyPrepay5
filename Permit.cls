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
Private ImpTim$
Private ImpDte As Date
Sub Import(P_Fx)
Fx = P_Fx
Stop
If IsEr Then
    WrtER
    MovToER
    Log MsgEr
Else
    RunSqlAy InsTbl_Permit_and_PermitD_SqlAy, WrkDb
    RmvWsEr
    MovToOK
    Log MsgOK
End If
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
DmpAy ErMsgAy
End Sub

Private Function ErMsgAy() As String()
Static Z As Boolean, O$()
If Not Z Then
    Z = True
    PushAy O, Er_MissingColumn: If Not IsEmptyAy(O) Then ErMsgAy = O: Exit Function
    PushAy O, Er_InvalidDtaTy:  If Not IsEmptyAy(O) Then ErMsgAy = O: Exit Function
    PushAy O, Er_DupSKU
    PushAy O, Er_ZerQty
    PushAy O, Er_NonIntQty
    PushAy O, Er_ZerAmt
End If
ErMsgAy = O
End Function

Private Function Er_ZerAmt() As String()
Const S$ = "Select PermitNo,Sku from Imp1 where Amt<=0"
Dim Dry: Dry = SqlDry(S, WrkDb1)
If IsEmptyAy(Dry) Then Exit Function
Dim O$()
Dim Dr
For Each Dr In Dry
    Push O, FmtQQ("PermitNo[?] Sku[?] has zero or negative Amount")
Next
Er_ZerAmt = O
End Function

Private Function WrkDb1() As Database
Static X As Database
If IsNothing(X) Then
    Set X = WrkDb
    RunSql "Select " & _
    "SKU," & _
    "[Material Description] As Des," & _
    "[Order Qty#]           As Qty," & _
    "[Permit Number]        As PermitNo," & _
    "[Batch Number]         As BchNo," & _
    "[Duty Paid Amount]     As Amt" & _
    " Into Imp1 from Imp", X
End If
Set WrkDb1 = X
End Function

Private Function WrkDb() As Database
Static X As Database
If IsNothing(X) Then
    Set X = TmpDb
    LnkWs Fx, "Sheet1", "Imp", X
End If
Set WrkDb = X
End Function

Sub Er_MissingColumn__Tst()
DmpAy Er_MissingColumn
End Sub

Private Function Er_MissingColumn() As String()
Dim Fny$(): Fny = TblFny("Imp", WrkDb)
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
Er_MissingColumn = O
End Function

Private Function Er_InvalidDtaTyMsg$(F$, ExpT As Dao.DataTypeEnum)
Dim T As Dao.DataTypeEnum: T = TblFld("Imp", F, WrkDb).Type
If T <> ExpT Then Er_InvalidDtaTyMsg = FmtQQ("Column[?] should be [?], but now it is[?]", F, DaoTyStr(T), DaoTyStr(ExpT))
End Function

Private Sub Er_InvalidDtaTy__Tst()
DmpAy Er_InvalidDtaTy
End Sub

Private Sub Er_Sheet1__Tst()
DmpAy Er_Sheet1
End Sub

Private Function Er_Sheet1() As String()
'Should only 1 sheet: Sheet1
'Or 2 sheets        : Sheet1 and Error
Dim WsNy$(): WsNy = FxWsNy(Fx)
Dim NWs%: NWs = Sz(WsNy)
Dim O$()
    If NWs <> 1 And NWs <> 2 Then
        Push O, FmtQQ("Excel should one or two worksheets, but now it is [?] worksheets", NWs)
        Er_Sheet1 = O
        Exit Function
    End If
    If NWs = 1 Then
        If WsNy(0) <> "Sheet1" Then
            Push O, "Worksheet name [Sheet1] not found"
        End If
    End If
    If NWs = 2 Then
        If (WsNy(0) <> "Error" And WsNy(0) <> "Sheet1") Or (WsNy(1) <> "Error" And WsNy(1) <> "Sheet1") Then
            Push O, FmtQQ("The 2 worksheet name should be [Sheet1] and [Error], but now they are [?] and [?]", WsNy(0), WsNy(1))
        End If
    End If
Er_Sheet1 = O
End Function

Private Function Er_InvalidDtaTy() As String()
Dim O$()
PushNonEmpty O, Er_InvalidDtaTyMsg("Material Description", dbText)
PushNonEmpty O, Er_InvalidDtaTyMsg("Order Qty#", dbDouble)
PushNonEmpty O, Er_InvalidDtaTyMsg("Permit Number", dbText)
PushNonEmpty O, Er_InvalidDtaTyMsg("Batch Number", dbText)
PushNonEmpty O, Er_InvalidDtaTyMsg("Duty Paid Amount", dbDouble)
Er_InvalidDtaTy = O
End Function

Private Sub Er_DupSKU__Tst()
DmpAy Er_DupSKU
End Sub

Private Function Er_DupSKU() As String()
Dim Dry(): Dry = SqlDry("Select Distinct PermitNo,SKU,Count(*) as Cnt from Imp1 Group By PermitNo,SKU Having Count(*)>1", WrkDb1)
Dim O$()
    Dim Dr
    If Not IsEmptyAy(Dry) Then
        Dim PermitNo$, Sku$, Cnt%
        For Each Dr In Dry
            PermitNo = Dr(0)
            Sku = Dr(1)
            Cnt = Dr(2)
            Push O, FmtQQ("There are [?] records with Same PermitNo[?] and SKU[?]", Cnt, PermitNo, Sku)
        Next
    End If
Er_DupSKU = O
End Function

Private Function Er_ZerQty() As String()
Const S$ = "Select PermitNo,Sku from Imp1 where Qty<=0"
Dim Dry: Dry = SqlDry(S, WrkDb1)
If IsEmptyAy(Dry) Then Exit Function
Dim O$()
Dim Dr
For Each Dr In Dry
    Push O, FmtQQ("PermitNo[?] Sku[?] has zero or negative quantity")
Next
Er_ZerQty = O
End Function

Private Function Er_NonIntQty() As String()
Const S$ = "Select PermitNo,Sku from Imp1 where Val(Qty)<>Qty"
Dim Dry: Dry = SqlDry(S, WrkDb1)
If IsEmptyAy(Dry) Then Exit Function
Dim O$()
Dim Dr
For Each Dr In Dry
    Push O, FmtQQ("PermitNo[?] Sku[?] has non-integer quantity")
Next
End Function

Private Function IsEr() As Boolean
IsEr = Not IsEmptyAy(ErMsgAy)
End Function

Private Sub MovToER()
Fso.MoveFile Fx, ErFfn
End Sub

Private Sub MovToOK()
Fso.MoveFile Fx, OkFfn
End Sub

Private Function OkPth$()
Dim O$: O = PermitImpPth & "OK\"
EnsPth O
OkPth = O
End Function

Private Function ErPth$()
Dim O$: O = PermitImpPth & "ER\"
EnsPth O
ErPth = O
End Function

Private Function OkFfn$()
Dim O$: O = OkPth & ImpTim & "\"
EnsPth O
OkFfn = O & Fx
End Function

Private Function ErFfn$()
Dim O$: O = ErPth & ImpTim & "\"
EnsPth O
ErFfn = O & Fx
End Function

Private Sub WrtER()
Dim Wb As Workbook: Set Wb = ImpPermWb
DltWs Wb, "Error"
Dim Ws As Worksheet: Set Ws = Wb.Sheets.Add
Ws.Name = "Error"
PutAy WsA1(Ws), ErMsgAy
Wb.Close True
End Sub

Private Function ImpPermHdrDrs() As Drs
'Summing following in WrkDb1->Imp1
Const Sql$ = "Select Distinct PermitNo,Count(*) as NSKU,Sum(x.Qty) as Qty,Sum(x.Amt) as Amt from Imp1 x Group By PermitNo"
ImpPermHdrDrs = SqlDrs(Sql, WrkDb1)
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
Function Ins_TblPermitD_SqlAy() As String()

End Function
Function Ins_TblPermit_SqlAy() As String()
Dim ODry()
    Dim M As PermitDftVal: M = PermitDftVal
    Dim Dr
    Dim ODr()
    Dim Dry()
    For Each Dr In ImpPermHdrDrs.Dry 'PermitNo NSKU Qty Amt
        Erase ODr
        Push ODr, DicOfId_ByPermitNo(Dr(0))
        Push ODr, Dr(0) ' PermitNo
        Push ODr, ImpDte
        Push ODr, ImpDte
        Push ODr, Dr(1) 'NSKU
        Push ODr, Dr(2) 'Qty
        Push ODr, Dr(3) 'Amt
        Push ODr, M.GLAc
        Push ODr, M.GLAcName
        Push ODr, M.BankCode
        Push ODr, M.ByUsr
        Push ODr, Now
        Push ODr, Empty
        Push ODr, True
        Push ODry, ODr
    Next
Dim Dt As Dt
    Dt.Fny = SplitSpc("Permit PermitNo PermitDate PostDate NSku Qty Tot GLAc GLAcName BankCode ByUsr DteCrt DteUpd IsImport")
    Dt.Dry = Dry
    Dt.DtNm = "Permit"
Ins_TblPermit_SqlAy = InsDtSqlAy(Dt)
End Function

Private Function InsTbl_Permit_and_PermitD_SqlAy() As String()
Dim O$()
Push O, Dlt_TblPermit_Sql
Push O, Dlt_TblPermitD_Sql
PushAy O, Ins_TblPermit_SqlAy
PushAy O, Ins_TblPermitD_SqlAy
End Function
Private Function DicOfId_ByPermitNo() As Dictionary

End Function
Private Function Dlt_TblPermit_Sql$()
Const S$ = "Delete From Permit Where PermitNo in (?)"
Dlt_TblPermit_Sql = FmtQQ(S, PermitNoLis)
End Function
Private Function Dlt_TblPermitD_Sql$()
Const S$ = "Delete From PermitD Where PermitNo in (?)"
Dlt_TblPermitD_Sql = FmtQQ(S, PermitNoLis)
End Function
Private Function PermitNoLis$()
PermitNoLis = JnComma(QuoteAy(PermitNoAy, "'"))
End Function
Private Function PermitNoAy() As String()
Static Z As Boolean, X$()
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
Function ImpPermWs() As Worksheet
Set ImpPermWs = ImpPermWb.Sheets(1)
End Function
Function ImpPermWb() As Workbook
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
    X = SelSq(Sq, MapStr)
End If
ImpPermDrs = X
End Function

Private Function MsgOK$()
MsgOK = "sdff"
End Function

Private Function MsgEr$()
End Function

Private Sub Class_Initialize()
ImpDte = Date
ImpTim = NowStr
End Sub
