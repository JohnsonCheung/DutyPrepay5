VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmYrAdj"
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

Private Sub CmdDetail_Click()
DoCmd.OpenForm "frmYrAdjD", OpenArgs:=Me.Year.Value - 2000
End Sub

Private Sub CmdRpt_Click()
Dim mY As Byte: mY = CByte(Me.Year.Value - 2000)
bCmd.CmdRpt mY
End Sub

Private Sub Form_Load()
DoCmd.Maximize
bYrO.AddYrO
Me.Requery
Me.Refresh
End Sub
