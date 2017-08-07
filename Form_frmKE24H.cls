VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmKE24H"
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

Private Sub Cmd_ShwDetail_Click()
Dim mA$: mA = Me.Year.Value - 2000 & "," & Me.Mth.Value
DoCmd.OpenForm "frmKE24", OpenArgs:=mA
End Sub

Private Sub CmdClear_Click()
bCmd.CmdKE24Clear CByte(Me.Year.Value - 2000), Me.Mth.Value
Me.Requery
Me.Recalc
Me.Refresh
End Sub

Private Sub CmdImport_Click()
bCmd.CmdKE24Import CByte(Me.Year.Value - 2000), Me.Mth.Value
Me.Requery
Me.Refresh
End Sub

Private Sub CmdOpnImportDir_Click()
bCmd.CmdOpnDirImport
End Sub

Private Sub Form_Open(Cancel As Integer)
'Aim: There is no current Yr record in table YrO, create one record in YrO
DoCmd.Maximize
Form_Open_1BldKE24H
End Sub

Private Sub Form_Open_1BldKE24H()
'AIm: From Min(Yr) of YrOD insert one record to KE24H if no such month record until current month
With CurrentDb.OpenRecordset("Select Min(Yr) from YrOD")
    Dim mYFm%
    mYFm = .Fields(0).Value
    .Close
End With
Dim J%
Dim mYTo%: mYTo = VBA.Year(Date)
mYTo = mYTo - 2000
For J = mYFm To mYTo
    Form_Open_1BldKE24H_1Y CByte(J)
Next
End Sub

Private Sub Form_Open_1BldKE24H_1Y(pY As Byte)
Dim J%
Dim mM As Byte: mM = IIf(pY = VBA.Year(Date) - 2000, Month(Date), 12)
For J = 1 To mM
    Form_Open_1BldKE24H_1Y_1M pY, CByte(J)
Next
End Sub

Private Sub Form_Open_1BldKE24H_1Y_1M(pY As Byte, pM As Byte)
With CurrentDb.OpenRecordset(Fmt_Str("Select Yr from KE24H where Yr={0} and Mth={1}", pY, pM))
    If .EOF Then .Close: DoCmd.RunSql Fmt_Str("Insert Into KE24H (Yr,Mth) values ({0},{1})", pY, pM): Exit Sub
    .Close
End With
End Sub

Private Sub Tot_DblClick(Cancel As Integer)
Dim mA$: mA = Me.Year.Value - 2000 & "," & Me.Mth.Value
DoCmd.OpenForm "frmKE24", OpenArgs:=mA
End Sub
