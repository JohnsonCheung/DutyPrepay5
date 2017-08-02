VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmBchRate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Option Base 0
Private xYY As Byte, xMM As Byte, xDD As Byte
Private Sub Cmd_AddToDummyPermit_Click()
If IsNull(Me.IsNoAssign.Value) Then Exit Sub
If Not Me.IsNoAssign.Value Then MsgBox "SKU[" & Me.Sku.Value & "] does not need to add to dummy permit", vbInformation: Exit Sub
Cmd_AddToDummyPermit_Click_1Add Me.Sku.Value
Cmd_AddToDummyPermit_Click_2Reset Me.Sku.Value
Cmd_AddToDummyPermit_Click_3UpdPermitThreeSum
Requery
End Sub
Private Sub Cmd_AddToDummyPermit_Click_1Add(pSku$)
Cmd_AddToDummyPermit_Click_1Add_1Permit
Cmd_AddToDummyPermit_Click_1Add_2PermitD pSku
End Sub
Private Sub Cmd_AddToDummyPermit_Click_1Add_1Permit()
With CurrentDb.OpenRecordset("Select * from Permit where Permit=1")
    If .EOF Then
        .Close
        DoCmd.RunSql "Insert into Permit (Permit,PermitNo,PermitDate,PostDate,GLAc,GLAcName,BankCode,ByUsr) values (1,'--Dummy--',#2000/1/1#,#2000/1/1#,'-','-','-','-')"
        Exit Sub
    End If
    .Close
End With
End Sub
Private Sub Cmd_AddToDummyPermit_Click_1Add_2PermitD(pSku$)
With CurrentDb.OpenRecordset(Fmt_Str("Select * from frmBchRate_OH where Sku='{0}' and IsNoAssign", pSku))
    While Not .EOF
        Cmd_AddToDummyPermit_Click_1Add_2PermitD_1Ins pSku, !BchNo, !OH
        .MoveNext
    Wend
    .Close
End With
End Sub
Private Sub Cmd_AddToDummyPermit_Click_1Add_2PermitD_1Ins(pSku$, pBchNo$, pQty&)
Dim mSeqNo%: mSeqNo = Nz(SqlToV("Select Max(SeqNo) from PermitD where Permit=1"), 0) + 10
Dim mRate@: mRate = Nz(SqlToV(Fmt_Str("Select DutyRateBott from frmBchRate where Sku='{0}'", pSku)), 0)
Dim mAmt@: mAmt = pQty * mRate
With CurrentDb.TableDefs("PermitD").OpenRecordset
    .AddNew
    !Permit = 1
    !Sku = pSku
    !SeqNo = mSeqNo
    !Qty = pQty
    !BchNo = pBchNo
    !Rate = mRate
    !Amt = mAmt
    .Update
    .Close
End With
End Sub
Private Sub Cmd_AddToDummyPermit_Click_2Reset(pSku$)
DoCmd.RunSql Fmt_Str("Update frmBchRate_OH set IsNoAssign=False where IsNoAssign and SKU='{0}'", pSku)
DoCmd.RunSql Fmt_Str("Update frmBchRate    set IsNoAssign=False where IsNoAssign and SKU='{0}'", pSku)
End Sub
Private Sub Cmd_AddToDummyPermit_Click_3UpdPermitThreeSum()
Dim mNSku%: mNSku = Nz(SqlToV("Select Count(*) from PermitD where Permit=1"), 0)
Dim mQty&: mQty = Nz(SqlToV("Select Sum(Qty) from PermitD where Permit=1"), 0)
Dim mTot@: mTot = Nz(SqlToV("Select Sum(Amt) from PermitD where Permit=1"), 0)
With CurrentDb.OpenRecordset("Select * from Permit where Permit=1")
    .Edit
    !DteUpd = Now()
    !NSku = mNSku
    !Qty = mQty
    !Tot = mTot
    .Update
End With
End Sub
Private Sub Cmd_Assign_Click()
bAssign.rr
zBldRecordSourceTable
End Sub

Private Sub CmdReadMe_Click()
MsgBox "Each latest inventory on hand SKU + batch# of taxable item in company 8600 at location Consignment & TaxPaid should assign to a permit line so that the inventory Duty-Rate of the inventory can be determined." & vbLf & vbLf & _
"Click [Assign Batch# to permit] will automatically assign batch# to permit." & vbLf & vbLf & _
"However some SKU may not able find any permit to assign.  In this case, [Cannot Assign] will be checked.  Click [Assign to dummy permit].  User can input the a Tax Duty in this dummy permit."
End Sub

Private Sub Form_Close()
OpnSkuBchRateErr
End Sub
Private Sub Form_Open(Cancel As Integer)
'Aim: Set xDteOH
'     Build "frmBchRate_OH = SKU OH from
DoCmd.Maximize
SetLatestYYMMDD xYY, xMM, xDD
Me.xDteOH.Value = "20" & Format(xYY, "00") & "-" & Format(xMM, "00") & "-" & Format(xDD, "00")
BuildSkuB
zBldRecordSourceTable
End Sub
Private Sub zBldRecordSourceTable()
'Aim: Refresh table-(frmBchRate frmBchRate_OH) from OH (Co=8600)
'Refresh frmBchRate = Sku DesSku OH IsNoAssign DutyRate BottPerAc DutyRateBott
'Refresh frmBchRate_OH = Sku BchNo OH IsNoAssign
DoCmd.SetWarnings False
'---- refresh table-frmBchRate_OH ------------------------------------------
DoCmd.RunSql "Delete from frmBchRate_OH"

DoCmd.RunSql Fmt_Str("Insert into frmBchRate_OH (Sku,BchNo,OH,IsNoAssign)" & _
" Select Distinct SKU, BchNo, Sum(Bott), False" & _
" from OH x inner join YpStk a on x.YpStk=a.YpStk" & _
" where IsTaxPaid and YY={0} and MM={1} and DD={2} and Co=1" & _
" and SKU in (Select SKU from SKU_StkHld where TaxRate is not null)" & _
" group by SKU,BchNo" & _
" having Sum(x.bott)<>0", xYY, xMM, xDD)    ' 'Co=1 is 8600 HK

DoCmd.RunSql "Select Sku,BchNo,Sum(Qty) as QtyPermit into `#frmBchRate_Permit` from PermitD where Nz(BchNo,'')<>'' Group by Sku,BchNo"
DoCmd.RunSql "Update frmBchRate_OH x left join `#frmBchRate_Permit` a on a.Sku=x.Sku and a.BchNo=x.BchNo set IsNoAssign=True where OH>Nz(QtyPermit,0)"
DoCmd.RunSql "Drop Table `#frmBchRate_Permit`"
'---- refresh table-frmBchRate ------------------------------------------
DoCmd.RunSql "Delete from frmBchRate"
DoCmd.RunSql "Insert into frmBchRate (Sku,OH,IsNoAssign) Select Distinct Sku, Sum(OH), Min(IsNoAssign) from frmBchRate_OH group by Sku"
DoCmd.RunSql "Update frmBchRate x inner join Sku_StkHld a on a.Sku=x.Sku set x.DutyRate=a.TaxRate"
DoCmd.RunSql "Update frmBchRate x inner join tblSKU a on a.SkuTxt=x.Sku set x.BottPerAc=a.BtlPerCs"
DoCmd.RunSql "Update (frmBchRate x inner join Sku88 a on a.Sku=x.Sku) inner join tblSku b on b.SkuTxt=a.Sku88 set x.BottPerAc=b.BtlPerCs"
DoCmd.RunSql "Update frmBchRate set DutyRateBott=DutyRate/BottPerAc where Nz(BottPerAc,0)<>0"
DoCmd.RunSql "Update frmBchRate x inner join qSku a on a.Sku=x.Sku set DesSku=`Sku Description`"
Me.Requery
Me.Refresh
End Sub
Private Sub Cmd_Close_Click()
DoCmd.Close
End Sub
