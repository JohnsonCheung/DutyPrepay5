VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmKE24"
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

Private Sub Form_Open(Cancel As Integer)
If VarType(Me.OpenArgs) <> vbString Then MsgBox "Me.OpenArgs is not a string.  It is supposed to be {Y},{M}": Cancel = True: Exit Sub
DoCmd.Maximize
Dim mA$(): mA = Split(Me.OpenArgs, ",")
Dim mY As Byte: mY = Val(mA(0))
Dim mM As Byte: mM = Val(mA(1))
Me.xYear.Value = mY + 2000
Me.xMth.Value = mM
Me.RecordSource = Fmt_Str("SELECT x.*, a.*, CCur(Tot/Qty) AS Rate" & _
" FROM KE24 x LEFT JOIN qSKU a ON x.Sku = a.Sku" & _
" Where Yr={0} and Mth={1}" & _
" Order by PostDate Desc, CopaNo Desc, CopaLNo Asc", mY, mM)
End Sub

