Attribute VB_Name = "bZZ"
Option Compare Database
Option Explicit
Type PermitDftVal
    GLAc As String
    GLAcName  As String
    ByUsr As String
    BankCode As String
End Type

Function PermitDftVal() As PermitDftVal
Dim O As PermitDftVal
With CurrentDb.OpenRecordset("Select * from Default")
    O.GLAc = !GLAc
    O.GLAcName = !GLAcName
    O.ByUsr = !ByUsr
    O.BankCode = !BankCode
    .Close
End With
PermitDftVal = O
End Function

Sub zCommon_CmdReadMe()
xOpn.Opn_ReadMe "DutyPrepay"
End Sub

Sub zFrmPermit_CmdDelete()
Form_frmPermit.CmdDelete_Click
End Sub

Function VdtBchNo(pBchNo$, pSku$, pRate@, ByRef oRate@) As Boolean
'Aim: Validate pBchNo:
'     - Return false for no error for pBchNo=''
'     - Return true for error for pSku='' or pRate=0
'     - If there is a record in SkuB return false. Set oRate=SkuB->Rate
'     - If there is no recor in SkuB return false. Add one record to SkuB, set oRate=pRate
If pBchNo = "" Then Exit Function
If pSku = "" Then MsgBox "SKU is blank": GoTo E
If pRate = 0 Then MsgBox "Rate is zero": GoTo E
With CurrentDb.OpenRecordset(Fmt_Str("Select * from SkuB where Sku='{0}' and BchNo='{1}'", pSku, pBchNo))
    If .EOF Then
        oRate = pRate
        .AddNew
        !DutyRateB = pRate
        !Sku = pSku
        !BchNo = pBchNo
        .Update
        .Close
        Exit Function
    End If
    oRate = !DutyRateB
    .Close
End With
Exit Function
E: VdtBchNo = True
End Function

Sub AddYrO()
'Aim: There is no current Yr record in table YrO, create one record in YrO
With CurrentDb.OpenRecordset("Select Yr from YrO where Yr=" & VBA.Year(Date) - 2000)
    If Not .EOF Then .Close: Exit Sub
    .Close
End With
DoCmd.RunSql "Insert Into YrO (Yr) values (Year(Date())-2000)"
End Sub

Function IsLasYrOD_Exist(pY As Byte) As Boolean
Dim mYr As Byte
mYr = pY - 1
With CurrentDb.OpenRecordset("Select count(*) from YrOD where Yr=" & mYr)
    If Nz(.Fields(0).Value, 0) > 0 Then IsLasYrOD_Exist = True
End With
End Function

Sub SetLatestYYMMDD(ByRef oYY As Byte, ByRef oMM As Byte, ByRef oDD As Byte)
Dim M$: M = SqlToV("Select Max(YY*10000+MM*100+DD) from OH")
oYY = Left(M, 2)
oMM = Mid(M, 3, 2)
oDD = Right(M, 2)
End Sub

Function CvPermit_NxtSeqNo%(pPermit&)
With CurrentDb.OpenRecordset("Select Max(SeqNo) from PermitD where Permit=" & pPermit)
    CvPermit_NxtSeqNo = 10 + Nz(.Fields(0).Value, 0)
    .Close
End With
End Function

Function FndRate_BySku(ByRef oRate@, pSku) As Boolean
With CurrentDb.OpenRecordset("Select DutyRate from qSku where Sku='" & pSku & "'")
    If .EOF Then .Close: MsgBox "SKU [" & pSku & "] not found", Buttons:=vbCritical: FndRate_BySku = True: Exit Function
    oRate = Nz(.Fields(0).Value, 0)
    .Close
End With
End Function

Function GetDesSku_BySku$(pSku)
If IsNull(pSku) Then Exit Function
With CurrentDb.OpenRecordset("Select `Sku Description` from qSku where Sku = '" & pSku & "'")
    If .EOF Then .Close: Exit Function
    GetDesSku_BySku = .Fields(0).Value
    .Close
End With
End Function

Function GetDirImport$()
GetDirImport = fct.CurMdbDir & "SAPDownloadExcel\"
End Function

Function GetLsFx_KE24$(pY As Byte, pM As Byte)
'Aim: Get list of Fx separated by VbLf in .\Import\KE24 yyyy-mm*.xls
Dim mDir$: mDir = fct.CurMdbDir & "SAPDownloadExcel\"
Dim mFxSpec$: mFxSpec = mDir & "KE24 " & pY + 2000 & "-" & Format(pM, "00") & "*.xls"
Dim mA$: mA = Dir(mFxSpec): If mA = "" Then MsgBox "No such file found:" & vbLf & vbLf & mFxSpec: Exit Function
mA = mDir & mA
Dim mB$: mB = Dir
While mB <> ""
    mA = mA & vbLf & mDir & mB
    mB = Dir
Wend
GetLsFx_KE24 = mA
End Function

Function GetNSku_YrOD%(pY As Byte)
With CurrentDb.OpenRecordset("Select Count(*) from YrOD where Yr=" & pY)
    If .EOF Then .Close: Exit Function
    GetNSku_YrOD = .Fields(0).Value
    .Close
End With
End Function

Function GetPermitNo_ByPermit$(pPermit&)
With CurrentDb.OpenRecordset("Select PermitNo from Permit where Permit=" & pPermit)
    If .EOF Then .Close: Exit Function
    GetPermitNo_ByPermit = .Fields(0).Value
    .Close
End With
End Function
Function SqlToAys(pSql$) As String()
Dim O$()
With CurrentDb.OpenRecordset(pSql)
    While Not .EOF
        If Not IsNull(.Fields(0).Value) Then Push O, .Fields(0).Value
        .MoveNext
    Wend
    .Close
End With
SqlToAys = O
End Function

Function SqlToInt%(pSql$)
With CurrentDb.OpenRecordset(pSql)
    SqlToInt = .Fields(0).Value
    .Close
End With
End Function

Function SqlToV(pSql$)
With CurrentDb.OpenRecordset(pSql)
    SqlToV = .Fields(0).Value
    .Close
End With
End Function
