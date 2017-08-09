Attribute VB_Name = "bSkuB"
Option Compare Database
Option Explicit

Sub BuildSkuB()
DoCmd.SetWarnings False
DoCmd.RunSql "Delete from SkuB"
DoCmd.RunSql "Insert into SkuB (Sku,BchNo,DutyRateB) select Distinct Sku,BchNo,Max(Rate) from PermitD where Nz(BchNo,'')<>'' group by Sku,BchNo"
End Sub
