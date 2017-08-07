VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmSKU88"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Option Base 0

Private Sub Cmd_Close_Click()
On Error GoTo Err_Cmd_Close_Click


    If Me.Dirty Then Me.Dirty = False
    DoCmd.Close

Exit_Cmd_Close_Click:
    Exit Sub

Err_Cmd_Close_Click:
    MsgBox Err.Description
    Resume Exit_Cmd_Close_Click
    
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
Me.DteUpd.Value = Now()
End Sub

Private Sub Form_Load()
DoCmd.Maximize
DoCmd.RunSql "INSERT INTO Sku88 (Sku88,Sku) SELECT SKUTxt,SKUTxt FROM tblSku x LEFT JOIN Sku88 a ON x.SKUTxt=a.Sku88 WHERE Left(SkuTxt,2)='88' AND Sku88 Is Null;"
Me.Requery
Me.Refresh
End Sub
