Attribute VB_Name = "bVdt"
Option Compare Database
Option Explicit

Function VdtBchNo(pBchNo$, pSku$, pRate@, ByRef oRate@) As Boolean
'Aim: Validate pBchNo:
'     - Return false for no error for pBchNo=''
'     - Return true for error for pSku='' or pRate=0
'     - If there is a record in SkuB return false. Set oRate=SkuB->Rate
'     - If there is no recor in SkuB return false. Add one record to SkuB, set oRate=pRate
If pBchNo = "" Then Exit Function
If pSku = "" Then MsgBox "SKU is blank": GoTo E
If pRate = 0 Then MsgBox "Rate is zero": GoTo E
With CurrentDb.OpenRecordset(Fmt_Str("Select * from SkuB where Sku='{0}' and BchNo='{1}'", pSku, pBchNo))
    If .EOF Then
        oRate = pRate
        .AddNew
        !DutyRateB = pRate
        !Sku = pSku
        !BchNo = pBchNo
        .Update
        .Close
        Exit Function
    End If
    oRate = !DutyRateB
    .Close
End With
Exit Function
E: VdtBchNo = True
End Function
