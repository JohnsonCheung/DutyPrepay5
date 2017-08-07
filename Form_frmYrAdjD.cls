VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmYrAdjD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Option Base 0

Private Sub AdjRate_AfterUpdate()
Me.NewRate.Value = Nz(Me.ClsRate.Value, 0) + Nz(Me.AdjRate.Value, 0)
Me.NewTot.Value = Nz(Me.NewRate.Value, 0) * Nz(Me.ClsQty.Value, 0)
Me.AdjTot.Value = Me.NewTot.Value - Nz(Me.ClsTot.Value, 0)
End Sub

Private Sub AdjTot_AfterUpdate()
Me.NewTot.Value = Nz(Me.ClsTot.Value, 0) + Nz(Me.AdjTot.Value, 0)
If Nz(Me.ClsQty.Value, 0) <> 0 Then
    Me.NewRate.Value = Nz(Me.NewTot.Value, 0) / Me.ClsQty.Value
Else
    Me.NewRate.Value = Null
End If
Me.AdjRate.Value = Nz(Me.NewRate.Value, 0) - Nz(Me.ClsRate.Value, 0)
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

Private Sub Form_AfterUpdate()
UpdTot
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
Me.DteUpd.Value = Now
If IsNull(Me.DteCrt.Value) Then Me.DteCrt.Value = Now()
End Sub

Private Sub Form_Close()
Form_frmYrAdj.Requery
End Sub

Private Sub Form_Close_1UpdYrAdjD()
'Aim: update YrAdjD by YrAdjDW
Dim mY As Byte: mY = Me.xYear.Value - 2000
DoCmd.RunSql Fmt_Str("SELECT {0} AS Yr, Sku, AdjTot, DteCrt, DteUpd INTO [#YrAdjD] FROM YrAdjDW WHERE AdjTot Is Not Null AND Round(AdjTot,2)<>0", mY)
        DoCmd.RunSql "DELETE FROM YrAdjD WHERE Yr=" & mY
        DoCmd.RunSql "INSERT INTO YrAdjD (Yr,Sku,AdjTot,DteCrt,DteUpd) SELECT Yr,Sku,AdjTot,DteCrt,DteUpd FROM [#YrAdjD];"
End Sub

Private Sub Form_Close_2UpdYrAdj()
'Aim: update YrAdj  by YrAdjD
Dim mY As Byte: mY = Me.xYear.Value - 2000
DoCmd.RunSql Fmt_Str("SELECT Yr, Count(1) AS NSku, Sum(x.AdjTot) AS AdjTot INTO [#YrO_FmYrAdjD] FROM YrAdjD x Where Yr={0} GROUP BY Yr;", mY)
DoCmd.RunSql "UPDATE YrO x INNER JOIN [#YrO_FmYrAdjD] a ON a.Yr=x.Yr SET x.AdjNSku=a.NSku, x.AdjTot=a.AdjTot, x.AdjDteUpd=Now() WHERE Nz(x.AdjNSku,0)<>a.NSku OR Nz(x.AdjTot,0)<>a.AdjTot;"
End Sub

Private Sub Form_Open(Cancel As Integer)
If VarType(Me.OpenArgs) <> vbString Then MsgBox "OpenArgs is not a string, which is supposed to be Yr", vbCritical: Cancel = True: Exit Sub
Me.xYear.Value = Me.OpenArgs + 2000
Me.RecordSource = "SELECT x.*, DutyRate, [Sku Description], [Business Area Code]" & _
" FROM YrAdjDW x LEFT JOIN qSku q ON x.Sku=q.Sku" & _
" Order by [Business Area Code],x.Sku"
Me.Requery
UpdTot
DoCmd.Maximize
End Sub

Private Sub Form_Open_1BldYrAdjDW(pY As Byte)
'Bld table-YrAdjDW"
DoCmd.RunSql "SELECT * INTO [#Mge] FROM YrAdjDW WHERE False"
DoCmd.RunSql "INSERT INTO [#Mge] (Sku,OpnQty,OpnTot,ClsQty,ClsTot) SELECT Sku, OpnQty,OpnTot,OpnQty,OpnTot FROM YrOD                                               WHERE Yr=" & pY
DoCmd.RunSql "INSERT INTO [#Mge] (Sku,ClsQty,ClsTot)               SELECT Sku, x.Qty, x.Amt                FROM PermitD x INNER JOIN Permit a ON a.Permit=x.Permit WHERE Year(PostDate)-2000)=" & pY
DoCmd.RunSql "INSERT INTO [#Mge] (Sku,ClsQty,ClsTot)               SELECT Sku, -Qty, -Tot                  FROM KE24                                               WHERE Yr=" & pY
DoCmd.RunSql "INSERT INTO [#Mge] (Sku,AdjTot,DteCrt,DteUpd)        SELECT Sku, AdjTot, DteCrt, DteUpd      FROM YrAdjD                                             WHERE Yr=" & pY

DoCmd.RunSql "SELECT Sku, Sum(x.OpnQty) AS OpnQty, Sum(x.OpnTot) AS OpnTot, CCur(0) AS OpnRate," & _
                        " Sum(x.ClsQty) AS ClsQty, Sum(x.ClsTot) AS ClsTot, CCur(0) AS ClsRate," & _
                                                 " Sum(x.AdjTot) AS AdjTot, CCur(0) AS AdjRate," & _
                                                 " Sum(x.NewTot) AS NewTot, CCur(0) AS NewRate," & _
                                                 " Max(x.DteCrt) AS DteCrt, Max(x.DteUpd) AS DteUpd" & _
" INTO [#Sum]" & _
" FROM [#Mge] x" & _
" GROUP BY Sku"

DoCmd.RunSql "DELETE FROM [#Sum] WHERE OpnQty=0 AND OpnTot=0 AND ClsQty=0 AND ClsTot=0 AND AdjTot=0;"
DoCmd.RunSql "UPDATE [#Sum] SET ClsRate = ClsTot/ClsQty WHERE ClsQty<>0;"
DoCmd.RunSql "UPDATE [#Sum] SET ClsRate = Null          WHERE Nz(ClsQty,0)=0"
DoCmd.RunSql "UPDATE [#Sum] SET OpnRate = OpnTot/OpnQty WHERE OpnQty<>0;"
DoCmd.RunSql "UPDATE [#Sum] SET NewTot = Nz(AdjTot,0)+Nz(ClsTot,0);"
DoCmd.RunSql "UPDATE [#Sum] SET NewRate= NewTot/ClsQty  WHERE .ClsQty<>0;"
DoCmd.RunSql "UPDATE [#Sum] SET AdjRate = IIf(NewRate=ClsRate,Null,NewRate-ClsRate);"
DoCmd.RunSql "DELETE FROM YrAdjDW;"
DoCmd.RunSql "INSERT INTO YrAdjDW SELECT * FROM [#Sum];"
End Sub

Private Sub NewRate_AfterUpdate()
Me.AdjRate.Value = Nz(Me.NewRate.Value, 0) - Nz(Me.ClsRate.Value, 0)
Me.NewTot.Value = Nz(Me.ClsQty.Value, 0) * Nz(Me.NewRate.Value, 0)
Me.AdjTot.Value = Nz(Me.NewTot.Value, 0) - Nz(Me.ClsTot.Value, 0)
End Sub

Private Sub NewRate_BeforeUpdate(Cancel As Integer)
If Nz(Me.NewRate.Value, 0) < 0 Then MsgBox "Cannot be -ve": Cancel = True
End Sub

Private Sub NewTot_AfterUpdate()
Me.AdjTot.Value = Nz(Me.NewTot.Value, 0) - Nz(Me.ClsTot.Value, 0)
If Nz(Me.ClsQty.Value, 0) <> 0 Then
    Me.NewRate.Value = Nz(Me.NewTot.Value, 0) / Me.ClsQty.Value
Else
    Me.NewRate.Value = Null
End If
Me.AdjRate.Value = Nz(Me.NewRate.Value, 0) - Nz(Me.ClsRate.Value, 0)
End Sub

Private Sub UpdAdjTot()
Dim mSql$
mSql = "Select" & _
", sum(AdjTot) as AdjTot" & _
", sum(NewTot) as NewTot" & _
", sum(IIf(Nz(AdjTot,0)=0,0,1)) as NSkuAdj" & _
" From YrAdjDW"
With CurrentDb.OpenRecordset(mSql)
    Me.xNSkuAdj.Value = Nz(!NSkuAdj, 0)
    Me.xAdjTot.Value = Nz(!AdjTot, 0)
    Me.xNewTot.Value = Nz(!NewTot, 0)
    .Close
End With
End Sub

Private Sub UpdTot()
Dim mSql$
mSql = "Select" & _
"  Count(*) as NSku" & _
", Sum(x.OpnQty) as OpnQty" & _
", Sum(x.OpnTot) as OpnTot" & _
", Sum(x.ClsQty) as ClsQty" & _
", Sum(x.ClsTot) as ClsTot" & _
", sum(x.AdjTot) as AdjTot" & _
", sum(x.NewTot) as NewTot" & _
", sum(IIf(Nz(x.AdjTot,0)=0,0,1)) as NSkuAdj" & _
" From YrAdjDW x"
With CurrentDb.OpenRecordset(mSql)
    Me.xNSku.Value = Nz(!NSku, 0)
    Me.xOpnQty.Value = Nz(!OpnQty, 0)
    Me.xOpnTot.Value = Nz(!OpnTot, 0)
    Me.xClsQty.Value = Nz(!ClsQty, 0)
    Me.xClsTot.Value = Nz(!ClsTot, 0)
    Me.xNSkuAdj.Value = Nz(!NSkuAdj, 0)
    Me.xAdjTot.Value = Nz(!AdjTot, 0)
    Me.xNewTot.Value = Nz(!NewTot, 0)
    .Close
End With
End Sub

