Attribute VB_Name = "bFrmPermitD"
Option Compare Database
Option Explicit
Function gChkRate$(pRate, pRateDuty)     ' Used by frmPermitD
If IsNull(pRateDuty) Then gChkRate = "No Rate!": Exit Function  ' pRateDuty is ZHT0 rate
If Nz(pRate, 0) = 0 Then gChkRate = "---": Exit Function
If pRate > pRateDuty Then
    If (pRate - pRateDuty) / pRate > 0.1 Then gChkRate = "Too low": Exit Function
Else
    If (pRateDuty - pRate) / pRate > 0.1 Then gChkRate = "Too high": Exit Function
End If
End Function
Function gDutyRateB@(pSku, pBchNo)     ' Used by frmPermitD
If IsNull(pSku) Then Exit Function         ' pRate     is PermitD->Rate which is user input
If IsNull(pBchNo) Then Exit Function     ' pRateDuty is ZHT0 rate
With CurrentDb.OpenRecordset(Fmt_Str("Select DutyRateB from SkuB where Sku='{0}' and BchNo='{1}'", pSku, pBchNo))
    If .EOF Then .Close: Exit Function
    gDutyRateB = Nz(.Fields(0).Value, 0)
    .Close
End With
End Function

