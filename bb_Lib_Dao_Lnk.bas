Attribute VB_Name = "bb_Lib_Dao_Lnk"
Option Compare Database
Option Explicit
Sub FxDb__Tst()
Dim Db As Database: Set Db = FxDb("N:\SapAccessReports\DutyPrepay5\SAPDownloadExcel\KE24 2010-01c.xls")
DmpAy Tny(Db)
Db.Close
End Sub
Function FxDb(Fx, Optional WsNmMapStr$) As Database
Dim O As Database
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
LnkWsAy Fx, WsNy, Tny, O
Set FxDb = O
End Function
Sub LnkWs(Fx, WsNm, Optional T, Optional D As Database)
Dim Db As Database: Set Db = DftDb(D)
Dim Tbl As Dao.TableDef
    Dim TblNm$: TblNm = Dft(T, WsNm)
    Set Tbl = Db.CreateTableDef(TblNm)
    Tbl.SourceTableName = WsNm & "$"
    Tbl.Connect = FmtQQ("Excel 8.0;HDR=YES;IMEX=2;DATABASE=?", Fx)
DrpTbl TblNm, Db
Db.TableDefs.Append Tbl
End Sub
Sub LnkWsAy(Fx, WsNy, Optional Tny, Optional D As Database)
Dim mTny: mTny = Dft(Tny, AddAyPfx(WsNy, ">"))
Dim Db As Database: Set Db = DftDb(D)
Dim J%
For J = 0 To UB(WsNy)
    LnkWs Fx, WsNy(J), mTny(J), Db
Next
End Sub
Function LnkTblFx$(T, Optional D As Database)
LnkTblFx = TakBet(Tbl(T, D).Connect, "Database=", ";")
End Function
