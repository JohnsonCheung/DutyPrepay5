VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmBchRate_Permit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Option Base 0

Private Sub BchNo_BeforeUpdate(Cancel As Integer)
Dim mRate@
If IsNull(Me.Sku.Value) Then MsgBox "SKU is empty": Cancel = True: Exit Sub
Cancel = BchNo_BeforeUpdate_1BchNo(Me.Sku.Value, Me.BchNo.Value): If Cancel Then Exit Sub
Cancel = VdtBchNo(Nz(Me.BchNo.Value), Nz(Me.Sku.Value, ""), Nz(Me.Rate.Value, 0), mRate)
End Sub

Private Function BchNo_BeforeUpdate_1BchNo(pSku$, pBchNo) As Boolean
'Aim: If pBchNo is null, return false for no error
'     If pSku+pBchNo not found in frmBchRate_OH, prompt message and return true for error
'Assume: pSku is non-blank
If IsNull(pBchNo) Then Exit Function
If Trim(pBchNo) = "" Then Exit Function
With CurrentDb.OpenRecordset(Fmt_Str("Select * from frmBchRate_OH where SKU='{0}' and BchNo='{1}'", pSku, pBchNo))
    If .EOF Then
        .Close
        MsgBox "The no such batch on hand!", vbCritical
        BchNo_BeforeUpdate_1BchNo = True
        Exit Function
    End If
    .Close
End With
End Function
