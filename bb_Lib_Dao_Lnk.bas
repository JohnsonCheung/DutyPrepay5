Attribute VB_Name = "bb_Lib_Dao_Lnk"
Option Compare Database
Option Explicit

Function FxDb(Fx, Optional WsNmMapStr$) As Db
Dim O As Db
    Set O = TmpDb
Dim WsNy$()
Dim Tny$()
    If WsNmMapStr = "" Then
        WsNy = FxWsNy(Fx)
    Else
        With BrkMapStr(WsNmMapStr)
            WsNy = .Sy2
            Tny = .Sy1
        End With
    End If
Db.LnkFx Fx, WsNy, Tny
Set FxDb = O
End Function


Private Sub FxDb__Tst()
Dim Db As Db: Set Db = FxDb("N:\SapAccessReports\DutyPrepay5\SAPDownloadExcel\KE24 2010-01c.xls")
AyDmp Db.Tny
Db.Db.Close
End Sub

Sub Tst()
FxDb__Tst
End Sub
