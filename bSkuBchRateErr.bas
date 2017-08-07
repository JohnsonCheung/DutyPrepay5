Attribute VB_Name = "bSkuBchRateErr"
Option Compare Database
Option Explicit

Sub OpnSkuBchRateErr()
'Aim: Build table SkuB and SkuBchRateErr
'     Open SkuBchRateErr if there is records
DoCmd.SetWarnings False
bSkuB.BuildSkuB

DoCmd.RunSql Fmt_Str("Select Distinct x.Sku,x.BchNo into `#SkuBchRateErr`" & _
" from (PermitD x" & _
" inner join SkuB a on a.Sku=x.Sku and a.BchNo=x.BchNo)" & _
" inner join Permit b on b.Permit=x.Permit" & _
" where Round(Rate,2)<>Round(DutyRateB,2)")

DoCmd.RunSql "Delete from SkuBchRateErr"
DoCmd.RunSql Fmt_Str("Insert into SkuBchRateErr" & _
         "(Permit,PermitNo,SeqNo,  Sku,  BchNo,Rate,DutyRateB)" & _
" Select a.Permit,PermitNo,SeqNo,x.Sku,x.BchNo,Rate,DutyRateB" & _
" from ((`#SkuBchRateErr` x" & _
" inner join PermitD a on a.Sku=x.Sku and a.BchNo=x.BchNo)" & _
" inner join Permit b on b.Permit=a.Permit)" & _
" inner join SkuB c on c.Sku=x.Sku and c.BchNo=x.BchNo")
DoCmd.RunSql "Update SkuBchRateErr set Diff = Nz(Rate,0)-Nz(DutyRateB,0) where Nz(Rate,0)<>Nz(DutyRateB,0)"
DoCmd.RunSql "Delete from SkuBchRateErr where Round(Diff,1)=0 or Diff is null"
DoCmd.RunSql "Update SkuBchRateErr x inner join Permit a on x.Permit=a.Permit Set x.PermitDate=a.PermitDate"
If CurrentDb.TableDefs("SkuBchRateErr").RecordCount > 0 Then DoCmd.OpenTable "SkuBchRateErr", acViewNormal, acReadOnly
End Sub

