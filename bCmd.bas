Attribute VB_Name = "bCmd"
Option Compare Database
Option Explicit
Option Base 0
Dim mXls As Excel.Application
Sub CmdGenFxPermit_Tst()
CmdGenFxPermit 10
End Sub
Sub CmdOpnPermitImpPth()
PthBrw PermitImpPth
End Sub
Sub CmdBldOpn(pY As Byte)
'Aim: Build Year Opening (table YrOd) by pY
'     Case#1: If not Last Year data, try import
'     Case#2: If there is last year data, build the YrOD and Update YrO
If VdtYr(pY) Then Exit Sub
If Not IsLasYrOD_Exist(pY) Then
    If Not Start("This is the first year opening." & vbLf & vbLf & "Import from Excel?") Then Exit Sub
    CmdBldOpn_1ImpFirstYrOpn pY
    Exit Sub
End If
If Not Start("Start building Year Opening [" & pY + 2000 & "]") Then Exit Sub
CmdBldOpn_2YrOD pY  ' Build current pY YrOD from last year YrOD+In-Out+Adj
CmdBldOpn_3YrO pY   ' Update YrO by YrOD of current pY
End Sub
Private Sub CmdBldOpn_1ImpFirstYrOpn(pY As Byte)
Dim mYear%: mYear = pY + 2000
Dim mFfn$: mFfn = GetDirImport & "Duty Prepay Year Opening " & mYear & ".xls"
If Dir(mFfn) = "" Then MsgBox "make sure this exists" & vbLf & mFfn: Exit Sub
If xCrt.Crt_Tbl_FmLnkWs(mFfn, "Sheet1", pNmtnew:=">FirstYrOpn") Then MsgBox "Cannot create link table [>FirstYrOpn]": Exit Sub
CmdBldOpn_1ImpFirstYrOpn_1Imp pY
CmdBldOpn_1ImpFirstYrOpn_2UpdYrO pY
End Sub
Private Sub CmdBldOpn_1ImpFirstYrOpn_1Imp(pY As Byte)
'Aim: Import >FirstYrOpn into YrOD
        DoCmd.RunSql "SELECT Trim(CStr(x.SKU)) AS SKU, Sum(x.Qty) AS Qty, Sum(x.Amt) as Amt INTO [#Inp] FROM [>FirstYrOpn] x GROUP BY Trim(CStr(SKU))"
        DoCmd.RunSql "DELETE FROM YrOD WHERE Yr=" & pY
DoCmd.RunSql Fmt_Str("INSERT INTO YrOD (Sku,OpnQty,OpnTot,Yr) SELECT SKU, Sum(Qty), Sum(Amt), {0} FROM [#Inp] GROUP BY SKU;", pY)
DoCmd.RunSql Fmt_Str("UPDATE YrOD SET OpnRate =OpnTot/OpnQty WHERE Yr={0};", pY)
End Sub
Private Sub CmdBldOpn_1ImpFirstYrOpn_2UpdYrO(pY As Byte)
DoCmd.RunSql Fmt_Str("SELECT Yr, Count(1) AS NSku, Sum(x.OpnQty) AS OpnQty, Sum(x.OpnTot) AS OpnTot INTO [#Tot] FROM YrOD x WHERE Yr={0} GROUP BY Yr;", pY)
DoCmd.RunSql "UPDATE YrO x INNER JOIN [#Tot] a ON a.Yr=x.Yr SET x.NSku=a.NSku, x.OpnQty=a.OpnQty, x.OpnTot=a.OpnTot, x.DteUpd= Now();"
End Sub
Private Sub CmdBldOpn_2YrOD(pY As Byte)
DoCmd.RunSql Fmt_Str("SELECT {0} AS Yr, Sku, OpnQty AS Q, OpnTot AS A INTO [#Mge]             FROM YrOD WHERE Yr={1}", pY, pY - 1)
DoCmd.RunSql Fmt_Str("INSERT INTO [#Mge] (Yr,Sku,Q,A) SELECT {0}, Sku, Sum(-Qty) , Sum(-Tot)  FROM KE24 WHERE Yr={1} GROUP BY Sku;", pY, pY - 1)
DoCmd.RunSql Fmt_Str("INSERT INTO [#Mge] (Yr,Sku,Q,A) SELECT {0}, Sku, Sum(x.Qty), Sum(x.Amt) FROM PermitD x INNER JOIN Permit ON x.Permit = a.Permit WHERE Year(PostDate)-2000)={1} GROUP BY Sku;", pY, pY - 1)
DoCmd.RunSql Fmt_Str("INSERT INTO [#Mge] (Yr,Sku,A)   SELECT {0}, Sku, AdjTot                 FROM YrAdjD WHERE Yr={1};", pY, pY - 1)

DoCmd.RunSql Fmt_Str("DELETE FROM YrOD WHERE Yr={0}", pY)
        DoCmd.RunSql "INSERT INTO YrOD (Yr,Sku,OpnQty,OpnTot) SELECT Yr, Sku, Sum(Q), Sum(A) FROM [#Mge] GROUP BY Yr,Sku;"
DoCmd.RunSql Fmt_Str("UPDATE YrOD SET OpnRate = OpnTot/OpnQty WHERE OpnQty<>0 AND Yr={0}", pY)
End Sub
Private Sub CmdBldOpn_3YrO(pY As Byte)
DoCmd.RunSql Fmt_Str("SELECT Yr, Count(1) AS NSku, Sum(x.OpnQty) AS OpnQty, Sum(x.OpnTot) AS OpnTot INTO [#Tot] FROM YrOD WHERE Yr={0} GROUP BY Yr;", pY)
DoCmd.RunSql "UPDATE YrO x INNER JOIN [#Tot] a ON a.Yr = x.Yr SET x.NSku=a.NSku, x.OpnQty = a.OpnQty, x.OpnTot = a.OpnTot, x.DteUpd = Now();"
End Sub
Private Sub CmdKE24Import__Tst()
CmdKE24Import 10, 1
End Sub
Sub CmdKE24Import(pY As Byte, pM As Byte)
If VdtYr(pY) Then Exit Sub
If VdtMth(pM) Then Exit Sub
Dim mLsFx$: mLsFx = GetLsFx_KE24(pY, pM): If mLsFx = "" Then Exit Sub
If Not Start("Following files are found, Import?" & vbLf & vbLf & mLsFx, "Import?") Then Exit Sub
Dim mAyFx$(): mAyFx = Split(mLsFx, vbLf)
Dim J%
DoCmd.SetWarnings False
For J = 0 To UBound(mAyFx)
    CmdKE24Import_1ImpFx mAyFx(J), pY, pM
Next
CmdKE24Import_2UpdKE24H pY, pM
Done
End Sub
Private Sub CmdKE24Import_1ImpFx(Pfx$, pY As Byte, pM As Byte)
'Aim: Import pFx$ into KE24 of pY, pM
If Dir(Pfx) = "" Then MsgBox "make sure this exists" & vbLf & Pfx: Exit Sub
If xCrt.Crt_Tbl_FmLnkWs(Pfx, "Sheet1", pNmtnew:=">KE24") Then MsgBox "Cannot create link table [>KE24] to Xls" & vbLf & Pfx: Exit Sub
With CurrentDb.OpenRecordset("Select Count(*) from [>KE24] where Year([Posting Date])<>" & 2000 + pY & " or Month([Posting Date])<>" & pM)
    If Nz(.Fields(0).Value, 0) <> 0 Then
        MsgBox "There are [" & .Fields(0).Value & "] records with Posting Date not in " & pY + 2000 & "/" & pM & vbLf & vbLf & Pfx, vbCritical, "Error in import file"
        .Close
        Exit Sub
    End If
End With
CmdKE24Import_1ImpKE24         ' Import >KE24 to KE24
CmdKE24Import_2UpdKE24H pY, pM ' Update current pY, pM of KE24H by KE24
End Sub
Private Sub CmdKE24Import_1ImpKE24()
DoCmd.RunSql "SELECT [Document number]                     AS CopaNo," & _
" CLng(IIf(Trim(Nz([Item number],''))='',0,[Item number])) AS CopaLNo," & _
                                          " [Posting date] AS PostDate," & _
                                 " CStr(Nz([Product],'-')) AS Sku," & _
                                  " CStr(Nz([Customer],0)) AS Cus," & _
                       " CLng(Nz(-[Billing qty in SKU],0)) AS Qty," & _
                              " CCur(Nz([D&T invoiced],0)) AS Tot" & _
" INTO [#KE24] FROM [>KE24];"
DoCmd.RunSql "INSERT INTO KE24 (CopaNo,   CopaLNo,   PostDate,   Sku,   Cus,   Qty,   Tot, Yr,                    Mth)" & _
                     " SELECT x.CopaNo, x.CopaLNo, x.PostDate, x.Sku, x.Cus, x.Qty, x.Tot, Year(x.PostDate)-2000, Month(x.PostDate)" & _
                     " FROM [#KE24] x " & _
                     " LEFT JOIN KE24 a ON x.CopaNo=a.CopaNo AND x.CopaLNo=a.CopaLNo" & _
                     " WHERE a.CopaNo Is Null;"
End Sub
Private Sub CmdKE24Import_2UpdKE24H(pY As Byte, pM As Byte)
DoCmd.RunSql Fmt_Str("SELECT Yr,Mth,Sum(x.Qty) AS Qty, Sum(x.Tot) AS Tot, Count(CopaLNo) AS NLin INTO [#Sum]        FROM KE24 x        WHERE Yr={0} And Mth={1} GROUP BY Yr,Mth;", pY, pM)
DoCmd.RunSql Fmt_Str("SELECT Yr,Mth, CopaNo                                                      INTO [#SumOrdList] FROM KE24          WHERE Yr={0} And Mth={1} GROUP BY Yr,Mth,CopaNo;", pY, pM)
        DoCmd.RunSql "SELECT Yr,Mth, Count(CopaNo) AS NOrd                                       INTO [#SumOrdCnt]  FROM [#SumOrdList]                          GROUP BY Yr,Mth;"
DoCmd.RunSql Fmt_Str("SELECT Yr,Mth, Cus                                                         INTO [#SumCusList] FROM KE24          WHERE Yr={0} And Mth={1} GROUP BY Yr,Mth,Cus;", pY, pM)
DoCmd.RunSql Fmt_Str("SELECT Yr,Mth, Count(Cus) AS NCus                                          INTO [#SumCusCnt]  FROM [#SumCusList]                          GROUP BY Yr,Mth;", pY, pM)
DoCmd.RunSql Fmt_Str("SELECT Yr,Mth, Sku                                                         INTO [#SumSkuList] FROM KE24          WHERE Yr={0} And Mth={1} GROUP BY Yr,Mth,Sku;", pY, pM)
DoCmd.RunSql Fmt_Str("SELECT Yr,Mth, Count(Sku) AS NSku                                          INTO [#SumSkuCnt]  FROM [#SumSkuList]                          GROUP BY Yr,Mth;", pY, pM)
DoCmd.RunSql Fmt_Str("UPDATE (((KE24H x" & _
                              " INNER JOIN [#Sum]       a ON x.Mth=a.Mth AND x.Yr=a.Yr)" & _
                              " INNER JOIN [#SumCusCnt] b ON x.Mth=b.Mth AND x.Yr=b.Yr)" & _
                              " INNER JOIN [#SumOrdCnt] c ON x.Mth=c.Mth AND x.Yr=c.Yr)" & _
                              " INNER JOIN [#SumSkuCnt] d ON x.Mth=d.Mth AND x.Yr=d.Yr" & _
                              " SET x.Qty     =a.Qty, x.Tot=a.Tot, x.NCopaLin=a.NLin," & _
                                  " x.NCus    =b.NCus," & _
                                  " x.NCopaOrd=c.NOrd," & _
                                  " x.NSku    =d.NSku, x.DteUpd = Now()" & _
                              " WHERE x.Qty     <>a.Qty Or x.Tot<>a.Tot Or x.NCopaLin<>a.NLin" & _
                                 " Or x.NCus    <>b.NCus" & _
                                 " Or x.NCopaOrd<>c.NOrd" & _
                                 " Or x.NSku    <>d.NSku" _
                                 , pY, pM)
End Sub
Sub CmdKE24Clear_Tst()
CmdKE24Clear 9, 2
End Sub
Sub CmdKE24Clear(pY As Byte, pM As Byte)
If VdtYr(pY) Then Exit Sub
If VdtMth(pM) Then Exit Sub
If Not Start("Clear sales history data (KE24) for Year[" & pY + 2000 & "] Month[" & pY & "]?", "Clear?") Then Exit Sub
Dim mCndn$: mCndn = Fmt_Str("Yr={0} and Mth={1}", pY, pM)
DoCmd.RunSql "Delete From KE24 where " & mCndn
DoCmd.RunSql "Update KE24H set NCopaOrd=0,NCopaLin=0,NCus=0,NSKU=0,Qty=0,Tot=0,DteUpd=Now() where " & mCndn
Done
End Sub
Private Sub CmdRpt__Tst()
CmdRpt 10
End Sub
Sub CmdRpt(pY As Byte)
'Aim Gen an Excel report with Month and Year by pY & pM
DoCmd.SetWarnings False
If pY > Year(Date) Then MsgBox "pY[" & pY & "] cannot > current year[" & Year(Date) & "].", vbCritical: Exit Sub
Dim mFxTo$: mFxTo = zzFxYrRpt(pY)
Dim mWb As Workbook
If Dir(mFxTo) <> "" Then
    If Not Start("Report exist, Regenerate?") Then
        xOpn.Opn_Wb mWb, mFxTo
        mWb.Application.Visible = True
        Exit Sub
    End If
End If

Dim mFxFm$: mFxFm = zzFxYrRptTp
Cpy_Fil mFxFm, mFxTo, pOvrWrt:=True
CmdRpt_1CrtOup pY ' Create @RptY & @RptM

Dim mWs As Worksheet
Opn_Wb mWb, mFxTo       ' The Tp contain query to @QryRptY & @QryRptM which will put additional columns to table @RptY & @RptM
xRfh.Rfh_Wb mWb
xSav.Sav_Wb mWb
mWb.Application.Visible = True
End Sub
Private Sub CmdRpt_1CrtOup__Tst()
DoCmd.SetWarnings False
CmdRpt_1CrtOup 10
End Sub
Private Sub CmdRpt_1CrtOup(pY As Byte)
'Aim: Create table @RptY @RptM from Permit,PermitD,KE24
'YpIO NmYpIO
'1   Opn
'2   In
'3   Out
'4   Close
'5   Adjusted
'6   New Clos
CmdRpt_1CrtOup_1Mge pY  ' Create #Mge
CmdRpt_1CrtOup_2M       ' Create @RptM from #Mge
CmdRpt_1CrtOup_3Y pY    ' Create @RptY
End Sub
Private Sub CmdRpt_1CrtOup_1Mge(pY As Byte)
CmdRpt_1CrtOup_1Mge_1YrOD pY       ' Create #Mge from YrOD
CmdRpt_1CrtOup_1Mge_2In pY         ' Insert #Mge from PermitD as In
CmdRpt_1CrtOup_1Mge_3Out pY        ' Insert #Mge from KE24    as Out
CmdRpt_1CrtOup_1Mge_4JanCls        ' Insert #Mge from Jan:Opn/In/Out as Jan:Cls
CmdRpt_1CrtOup_1Mge_5OpnCls pY     ' Insert #Mge from Feb-Dec:Opn/Cls
CmdRpt_1CrtOup_1Mge_6AdjYrD pY     ' Insert #Mge from AdjYrD
CmdRpt_1CrtOup_1Mge_7NewCls        ' Insert #Mge from #Mge for NewCls
DoCmd.RunSql "Delete from [#Mge] where Nz(A,0)=0 and Nz(Q,0)=0"
DoCmd.RunSql "UPDATE [#Mge] SET A=Null WHERE A=0;"
DoCmd.RunSql "UPDATE [#Mge] SET Q=Null WHERE Q=0;"
CmdRpt_1CrtOup_1Mge_8AllYpIO pY
End Sub
Private Sub CmdRpt_1CrtOup_1Mge_1YrOD(pY As Byte)
DoCmd.RunSql Fmt_Str("SELECT Yr, CByte(1) AS Mth, Sku, CByte(1) AS YpIO, Sum(OpnQty) AS Q, Sum(OpnTot) AS A INTO [#Mge] FROM YrOD Where Yr={0} GROUP BY Yr,Sku", pY)
End Sub
Private Sub CmdRpt_1CrtOup_1Mge_2In(pY As Byte)
DoCmd.RunSql Fmt_Str("INSERT INTO `#Mge` (Yr,Mth,Sku,YpIO,Q,A)" & _
" SELECT Year(PostDate)-2000, Month(PostDate), Sku, 2, Sum(a.Qty), Sum(a.Amt)" & _
" FROM Permit x INNER JOIN PermitD a ON a.Permit = x.Permit" & _
" Where Year(PostDate)-2000={0}" & _
" GROUP BY Year(PostDate)-2000, Month(PostDate), Sku", pY)
End Sub
Private Sub CmdRpt_1CrtOup_1Mge_3Out(pY As Byte)
DoCmd.RunSql Fmt_Str("INSERT INTO [#Mge] (Yr,Mth,Sku,YpIO,Q,A) SELECT Yr,Mth,Sku,3,Sum(Qty),Sum(Tot) FROM KE24 x WHERE Yr={0} GROUP BY Yr,Mth,Sku", pY)
End Sub
Private Sub CmdRpt_1CrtOup_1Mge_4JanCls()
DoCmd.RunSql "INSERT INTO [#Mge] (Yr,Mth,Sku,YpIO,Q,A) SELECT Yr,1,Sku,4,Sum(Q),Sum(A) FROM [#Mge] WHERE Mth=1 GROUP BY Yr,Sku;"
End Sub
Private Sub CmdRpt_1CrtOup_1Mge_5OpnCls(pY As Byte)
Dim J%
For J = 2 To zM(pY) ' If pY is current year, return current month else return 12
    DoCmd.RunSql Fmt_Str("INSERT INTO [#Mge] (Yr,Mth,Sku,YpIO,Q,A) SELECT Yr,{0},Sku,1,Sum(Q),Sum(A) FROM [#Mge] WHERE Mth={1} and YpIO=4 GROUP BY Yr,Sku;", J, J - 1)
    DoCmd.RunSql Fmt_Str("INSERT INTO [#Mge] (Yr,Mth,Sku,YpIO,Q,A) SELECT Yr,{0},Sku,4,Sum(Q),Sum(A) FROM [#Mge] WHERE Mth={0}            GROUP BY Yr,Sku;", J)
Next
End Sub
Private Sub CmdRpt_1CrtOup_1Mge_6AdjYrD(pY As Byte)
'Aim: Insert #Mge from AdjYrD
DoCmd.RunSql Fmt_Str("INSERT INTO [#Mge] (Yr,Mth,Sku,YpIO,A) SELECT Yr,13,Sku,5,Sum(AdjTot) FROM YrAdjD WHERE Yr={0} GROUP BY Yr,Sku;", pY)
End Sub
Private Sub CmdRpt_1CrtOup_1Mge_7NewCls()
DoCmd.RunSql "INSERT INTO [#Mge] (Yr,Mth,Sku,YpIO,Q,A) SELECT Yr,12,Sku,6,Sum(Q),Sum(A) FROM [#Mge] WHERE YpIO In (4,5) AND Mth=12 GROUP BY Yr,Sku;"
End Sub
Private Sub CmdRpt_1CrtOup_1Mge_8AllYpIO(pY As Byte)
'Aim: Write record to #Mge for ( 1st SKU x 12 month x 6 YpIO )
Dim mSku$: mSku = CmdRpt_1CrtOup_1Mge_8AllYpIO_1MinSku
Dim J%, I%
With CurrentDb.TableDefs("#Mge").OpenRecordset
    For J = 1 To zM(pY)
        For I = 1 To 6
            .AddNew
            !Yr = pY
            !Mth = J
            !Sku = mSku
            !YpIO = I
            !Q = 0
            !A = 0
            .Update
        Next
    Next
    .Close
End With
End Sub
Private Function CmdRpt_1CrtOup_1Mge_8AllYpIO_1MinSku$()
With CurrentDb.OpenRecordset("Select Min(Sku) from `#Mge`")
    CmdRpt_1CrtOup_1Mge_8AllYpIO_1MinSku = .Fields(0).Value
    .Close
End With
End Function
Private Sub CmdRpt_1CrtOup_2M()
DoCmd.RunSql "Delete from `@RptM`"
DoCmd.RunSql "Insert into `@RptM` Select * from `#Mge`"
End Sub
Private Sub CmdRpt_1CrtOup_3Y(pY As Byte)
CmdRpt_1CrtOup_3Y_1MgeIO
CmdRpt_1CrtOup_3Y_2MaxClsMth pY
DoCmd.RunSql "DELETE FROM [@RptY] WHERE Yr=" & pY
DoCmd.RunSql "INSERT INTO [@RptY] (Yr,Sku,YpIO,Q,A) SELECT   Yr,  Sku,  YpIO,  Q,  A FROM [@RptM]  WHERE YpIO=1 AND Mth=1 and Yr=" & pY
DoCmd.RunSql "INSERT INTO [@RptY] (Yr,Sku,YpIO,Q,A) SELECT   Yr,  Sku,  YpIO,  Q,  A FROM [#MgeIO] WHERE Yr=" & pY
DoCmd.RunSql "INSERT INTO [@RptY] (Yr,Sku,YpIO,Q,A) SELECT x.Yr,x.Sku,x.YpIO,x.Q,x.A FROM [@RptM] x INNER JOIN [#MaxClsMth] a ON x.Yr=a.Yr AND x.Mth=a.MaxClsMth AND x.YpIO=a.YpIO WHERE x.Yr=" & pY
DoCmd.RunSql "INSERT INTO [@RptY] (Yr,Sku,YpIO,Q,A) SELECT   Yr,  Sku,  YpIO,  Q,  A FROM [@RptM]  WHERE YpIO In (5,6) and Yr=" & pY
End Sub
Private Sub CmdRpt_1CrtOup_3Y_1MgeIO()
DoCmd.RunSql "SELECT Yr,Sku,YpIO,Sum(x.Q) AS Q, Sum(x.A) AS A INTO [#MgeIO] FROM [#Mge] x WHERE YpIO In (2,3) GROUP BY Yr,Sku,YpIO;"
End Sub
Private Sub CmdRpt_1CrtOup_3Y_2MaxClsMth(pY As Byte)
DoCmd.RunSql Fmt_Str("SELECT Yr,YpIO,Max(Mth) AS MaxClsMth INTO [#MaxClsMth] FROM [@RptM] WHERE Yr={0} And YpIO=4 GROUP BY Yr,YpIO;", pY)
End Sub
Sub CmdGenFxPermit(pPermit&)
'Aim Gen a Permit Excel Form by pPermit
DoCmd.SetWarnings False
Dim mFxTo$: mFxTo = zzFxPermit(pPermit)
Dim mWb As Workbook
If Dir(mFxTo) <> "" Then
    If Not Start("Form exist, Regenerate?") Then
        xOpn.Opn_Wb mWb, mFxTo
        mWb.Application.Visible = True
        Exit Sub
    End If
End If

Dim mFxFm$: mFxFm = zzFxTp

Cpy_Fil mFxFm, mFxTo, pOvrWrt:=True
'
Dim mWs As Worksheet
Opn_Wb mWb, mFxTo
Set mWs = mWb.Sheets(1)
'-- Fill in the mWs

'' Get Variables
'''Line: {GLAc} {GLAcName} {TxAmt}
'''{TotCr} {Tot} {PermitNo} {PermitId} {PostDate} {PermitDate} {BankCode} {ByUsr}
Dim xTotCr@, xTot@, xPermitNo$, xPermitId$, xPostDate$, xPermitDate$, xBankCode$, xByUsr$
Dim xGLAc$, xGLAcName$, xTxAmt@(), xBusArea$()
With CurrentDb.OpenRecordset("Select * from Permit where Permit=" & pPermit)
    If .EOF Then .Close: MsgBox "Cannot get record from table Permit by pPermit=[" & pPermit & "]": Exit Sub
    xTot = !Tot
    xTotCr = -xTot
    xPermitNo = !PermitNo
    xPermitId = Format(!Permit, "00000")
    xPostDate = Format(!PostDate, "yyyy-mm-dd")
    xPermitDate = Format(!PermitDate, "yyyy-mm-dd")
    xBankCode = Nz(!BankCode.Value, "")
    xByUsr = Nz(!ByUsr.Value, "")
    xGLAc = Nz(!GLAc.Value, "")
    xGLAcName = Nz(!GLAcName.Value, "")
    .Close
End With
Dim N%: N = 0
Dim mSql$
mSql = "SELECT [Business Area Code] as BusArea, Sum(x.Amt) as TxAmt" & _
" FROM PermitD x" & _
" INNER JOIN qSKU s ON x.Sku = s.Sku" & _
" WHERE x.Permit = " & pPermit & _
" GROUP BY [Business Area Code];"
With CurrentDb.OpenRecordset(mSql)
    If .EOF Then .Close: MsgBox "Cannot get record from table PermitD by pPermit=[" & pPermit & "]": Exit Sub
    While Not .EOF
        ReDim Preserve xTxAmt(N), xBusArea(N)
        xTxAmt(N) = 0 - Nz(!TxAmt, 0)
        xBusArea(N) = Nz(!BusArea, "")
        N = N + 1
        .MoveNext
    Wend
    .Close
End With

'' Fill in Ws by Variables
Dim mRge As Range
Set mRge = mWb.Names("PrintArea").RefersToRange
Dim mRnoBeg& ' The row with {BusArea}
Dim mCnoBusArea ' The column with {BusArea}
Dim mCnoTxAmt   ' The column with {TxAmt}
Dim iCell As Range
For Each iCell In mRge
    Dim mV: mV = iCell.Value
    If VarType(mV) = vbString Then
        Dim mS$: mS = mV
        If Left(mS, 1) = "{" Then
            Select Case mS
            Case "{Tot}": iCell.Value = xTot
            Case "{TotCr}": iCell.Value = xTotCr
            Case "{PermitNo}": iCell.Value = xPermitNo
            Case "{PermitId}": iCell.Value = xPermitId
            Case "{PostDate}": iCell.Value = xPostDate
            Case "{PermitDate}": iCell.Value = xPermitDate
            Case "{BankCode}": iCell.Value = xBankCode
            Case "{ByUsr}": iCell.Value = xByUsr
            Case "{GLAc}": iCell.Value = xGLAc
            Case "{GLAcName}": iCell.Value = xGLAcName
            Case "{BusArea}": mRnoBeg = iCell.Row: mCnoBusArea = iCell.Column
            Case "{TxAmt}": mCnoTxAmt = iCell.Column
            Case "{TimeStamp}": iCell.Value = Format(Now, "yyyy-mm-dd hh:nn")
            End Select
        End If
    End If
Next
'' Fill in Ws by TxAmt(), BusArea(), mRnoBeg, mCnoBusArea, mCnoTxAmt
If mRnoBeg = 0 Then
    MsgBox "No {BusArea} is found in the Template!!"
Else
    Dim J%
    Dim mRgeNxt As Range
    For J = 1 To UBound(xTxAmt)
        Set mRge = mWs.Rows(mRnoBeg)
        mRge.EntireRow.Select
        Selection.Copy
        Set mRgeNxt = mWs.Rows(mRnoBeg + 1)
        mRgeNxt.EntireRow.Select
        mWs.Paste
    Next
    For J = 0 To UBound(xTxAmt)
        Set mRge = mWs.Cells(J + mRnoBeg, mCnoTxAmt)
        mRge.Value = xTxAmt(J)
        Set mRge = mWs.Cells(J + mRnoBeg, mCnoBusArea)
        mRge.Value = xBusArea(J)
    Next
End If
DoCmd.RunSql "SELECT x.Sku, qSKU.[SKU Description], x.Amt, x.Rate, x.Qty INTO [@Permit]" & _
" FROM Permit AS a INNER JOIN (PermitD AS x LEFT JOIN qSKU ON x.Sku = qSKU.Sku) ON a.Permit = x.Permit" & _
" WHERE x.Permit = " & pPermit & _
" ORDER BY x.SeqNo;"
xRfh.Rfh_Wb mWb
xSav.Sav_Wb mWb
mWb.Application.Visible = True
End Sub
Sub CmdOpnDir()
xOpn.Opn_Dir fct.CurMdbDir & "Cheque Request\"
End Sub
Sub CmdOpnDirImport()
Dim mDir$: mDir = fct.CurMdbDir & "SAPDownloadExcel\"
xOpn.Opn_Dir mDir
End Sub
Private Function zM(pY As Byte) As Byte
If pY + 2000 = Year(Date) Then
    zM = Month(Date)
Else
    zM = 12
End If
End Function
Private Function zzFxTp$()
zzFxTp = fct.CurMdbDir & "Template\Template_DutyPrepay_Cheque_Request_Form.xls"
End Function
Private Function zzFxPermit(pPermit&)
Dim mFn$
Dim mPermitNo$: mPermitNo = GetPermitNo_ByPermit(pPermit)
If mPermitNo = "" Then MsgBox "gPermit() is 0": Exit Function
mFn = "(" & Format(pPermit, "00000") & ") " & mPermitNo & ".xls"
zzFxPermit = fct.CurMdbDir & "Cheque Request\" & mFn
End Function
Private Function zzFxYrRptTp$()
zzFxYrRptTp = fct.CurMdbDir & "Template\Template_DutyPrepay_Report.xls"
End Function
Private Function zzFxYrRpt$(pY As Byte)
zzFxYrRpt = fct.CurMdbDir & "Output\Duty prepay report - Year " & pY + 2000 & ".xls"
End Function

Sub Tst()
CmdKE24Import__Tst
CmdRpt__Tst
CmdRpt_1CrtOup__Tst
End Sub
