Attribute VB_Name = "g"
Option Compare Database
Option Base 0
Option Explicit
Type PermitDftVal
    GLAc As String
    GLAcName  As String
    ByUsr As String
    BankCode As String
End Type

Sub A()
Const cPfx$ = "qryRpt"
Dim J%
For J = 0 To CurrentDb.QueryDefs.Count - 1
    If Left(CurrentDb.QueryDefs(J).Name, Len(cPfx)) = cPfx Then Debug.Print CurrentDb.QueryDefs(J).Sql & vbLf
Next
End Sub

Property Get PermitDftVal() As PermitDftVal
Dim O As PermitDftVal
With CurrentDb.OpenRecordset("Select * from Default")
    O.GLAc = !GLAc
    O.GLAcName = !GLAcName
    O.ByUsr = !ByUsr
    O.BankCode = !BankCode
    .Close
End With
PermitDftVal = O
End Property

Function VdtMth(pM As Byte) As Boolean
If pM > 12 Or pM < 1 Then MsgBox "pM must between 1 and 12": VdtMth = True
End Function

Function VdtYr(pY As Byte) As Boolean
If pY = 0 Then MsgBox "pY is 0": VdtYr = True
End Function

Sub zCommon_CmdReadMe()
xOpn.Opn_ReadMe "DutyPrepay"
End Sub

Sub zFrmPermit_CmdDelete()
Form_frmPermit.CmdDelete_Click
End Sub
