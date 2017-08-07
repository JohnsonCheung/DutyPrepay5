VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmYrO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Option Base 0
Const cMod$ = cLib & ".Form_frmYrO"

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

Private Sub CmdBldOpn_Click()
If IsNull(Me.Yr.Value) Then Exit Sub
bCmd.CmdBldOpn CByte(Me.Yr.Value)
Me.Requery
End Sub

Private Sub CmdDetail_Click()
If IsNull(Me.Yr.Value) Then Exit Sub
DoCmd.OpenForm "frmYrOD", OpenArgs:=CByte(Me.Yr.Value)
End Sub

Private Sub CmdOpnImportDir_Click()
bCmd.CmdOpnDirImport
End Sub

Private Sub CmdRpt_Click()
If IsNull(Me.Yr.Value) Then Exit Sub
bCmd.CmdRpt Me.Yr.Value
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
Me.DteUpd.Value = Now
End Sub

Private Sub Form_Open(Cancel As Integer)
bYrO.AddYrO
DoCmd.Maximize
Me.Requery
Me.Recalc
Me.Refresh
End Sub
