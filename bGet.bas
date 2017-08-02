Attribute VB_Name = "bGet"
Option Compare Database
Option Base 0
Option Explicit
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
Function GetPermitNo_ByPermit$(pPermit&)
With CurrentDb.OpenRecordset("Select PermitNo from Permit where Permit=" & pPermit)
    If .EOF Then .Close: Exit Function
    GetPermitNo_ByPermit = .Fields(0).Value
    .Close
End With
End Function
Function GetNSku_YrOD%(pY As Byte)
With CurrentDb.OpenRecordset("Select Count(*) from YrOD where Yr=" & pY)
    If .EOF Then .Close: Exit Function
    GetNSku_YrOD = .Fields(0).Value
    .Close
End With
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

