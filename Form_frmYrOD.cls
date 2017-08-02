VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmYrOD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Option Base 0
Private Sub Form_Open(Cancel As Integer)
If VarType(Me.OpenArgs) <> vbByte Then MsgBox "Me.OpenArgs is not a byte!": Cancel = True: Exit Sub
DoCmd.Maximize
Me.xYear.Value = Me.OpenArgs + 2000
Me.RecordSource = "SELECT [Business Area Code], DutyRate, X.Sku, [Sku Description], OpnQty, OpnRate, OpnTot" & _
" FROM YrOD X left join qSku a on x.Sku = a.Sku" & _
" WHERE X.Yr = " & Me.OpenArgs & _
" ORDER BY [Business Area Code], x.Sku;"
UpdTot Me.OpenArgs
Me.Refresh
End Sub
Private Sub UpdTot(pY As Byte)
Dim mSql$
mSql = "Select" & _
"  Count(*) as NSku" & _
", Sum(x.OpnQty) as OpnQty" & _
", Sum(x.OpnTot) as OpnTot" & _
" From YrOD x where Yr=" & pY
With CurrentDb.OpenRecordset(mSql)
    Me.xNSku.Value = Nz(!NSku, 0)
    Me.xOpnQty.Value = Nz(!OpnQty, 0)
    Me.xOpnTot.Value = Nz(!OpnTot, 0)
    .Close
End With
End Sub

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
