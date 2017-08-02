VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmPermit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Option Base 0
Const cMod$ = cLib & ".Form_frmPermit"
Private xGLAc$, xGLAcName$, xByUsr$, xBankCode$
Private Sub CmdEdt_Click()
Sav_Rec
If IsNull(Me.Permit.Value) Then Exit Sub
DoCmd.OpenForm "frmPermitD", acNormal, DataMode:=acFormEdit, OpenArgs:=Me.Permit.Value
End Sub
Public Sub CmdDelete_Click()
Sav_Rec
If IsNull(Me.Permit.Value) Then Exit Sub
Dim mA$: mA = InputBox("Input Yes to delete permit[" & Me.PermitNo.Value & "]")
If mA <> "YES" Then Exit Sub
DoCmd.RunSql "Delete from Permit where Permit=" & Me.Permit.Value
Me.Requery
End Sub

Private Sub CmdExp_Click()
bFrmPermit.Exp
End Sub

Private Sub CmdOpnImpFdr_Click()
BrwPth PermitImpPth
End Sub

Private Sub CmdPrint_Click()
Sav_Rec
If IsNull(Me.Permit.Value) Then MsgBox "Please enter Permit#": Exit Sub
CmdGenFxPermit Me.Permit.Value
End Sub
Private Sub Form_BeforeInsert(Cancel As Integer)
Me.GLAc.Value = xGLAc
Me.GLAcName.Value = xGLAcName
Me.ByUsr.Value = xByUsr
Me.BankCode.Value = xBankCode
End Sub
Private Sub Form_BeforeUpdate(Cancel As Integer)
Me.DteUpd.Value = Now
End Sub
Private Sub Form_Open(Cancel As Integer)
DoCmd.Maximize
With PermitDftVal
    xGLAc = .GLAc
    xGLAcName = .GLAcName
    xByUsr = .ByUsr
    xBankCode = .BankCode
End With
BuildSkuB
DoCmd.RunCommand acCmdRemoveFilterSort
End Sub

Private Sub PermitNo_LostFocus()
Dim A$: A = Replace(Me.PermitNo.Value, Chr(9), "")
If A <> Me.PermitNo.Value Then
    Me.PermitNo.Value = A
End If
End Sub

Private Sub PostDate_BeforeUpdate(Cancel As Integer)
If IsNull(Me.PostDate.Value) Then MsgBox "Cannot be blank": Cancel = True: Exit Sub
If Me.PostDate.Value < #1/1/2000# Then MsgBox "Cannot less then 2010/01/01": Cancel = True: Exit Sub
If Me.PostDate.Value > #1/1/2050# Then MsgBox "Cannot greater then 2050/01/01": Cancel = True: Exit Sub
If VBA.Year(Me.PostDate.Value) <> VBA.Year(Date) Then If MsgBox("The year is not current year, is it OK?", vbOKCancel) = vbCancel Then Cancel = True
End Sub
Private Sub PermitDate_BeforeUpdate(Cancel As Integer)
If IsNull(Me.PermitDate.Value) Then MsgBox "Cannot be blank": Cancel = True: Exit Sub
If Me.PermitDate.Value < #1/1/2010# Then MsgBox "Cannot less then 2010/01/01": Cancel = True: Exit Sub
If Me.PermitDate.Value > #1/1/2050# Then MsgBox "Cannot greater then 2050/01/01": Cancel = True: Exit Sub
If VBA.Year(Me.PermitDate.Value) <> VBA.Year(Date) Then If MsgBox("The year is not current year, is it OK?", vbOKCancel) = vbCancel Then Cancel = True
End Sub
Private Sub Cmd_Close_Click()
DoCmd.Close
End Sub
