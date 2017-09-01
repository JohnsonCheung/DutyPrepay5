Attribute VB_Name = "Dao_Lnk"
Option Compare Database
Option Explicit

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
LnkFx Fx, WsNy, Tny, O
Set FxDb = O
End Function

Sub LnkFb(Fb, TblNm_or_Ny, Optional Tar, Optional D As Database)
Dim Src$(): Src = NyCv(TblNm_or_Ny)
Dim mTar$(): mTar = NyCv(Dft(Tar, Src))
Dim Db As Database: Set Db = DftDb(D)
Dim Tbl As Dao.TableDef
Dim J%
For J = 0 To UB(mTar)
    Set Tbl = Db.CreateTableDef(mTar(J))
    Tbl.SourceTableName = Src(J)
    Tbl.Connect = FmtQQ(";DATABASE=?", Fb)
    DrpTbl mTar(J), Db
    Db.TableDefs.Append Tbl
Next
End Sub

Sub LnkFx(Fx, WsNm_or_Ny, Optional Tar, Optional D As Database)
Dim Src$(): Src = NyCv(WsNm_or_Ny)
Dim mTar$(): mTar = NyCv(Dft(Tar, Src))
Dim Db As Database: Set Db = DftDb(D)
Dim Tbl As Dao.TableDef
Dim J%
For J = 0 To UB(mTar)
    Set Tbl = Db.CreateTableDef(mTar(J))
    Tbl.SourceTableName = Src(J) & "$"
    Tbl.Connect = FmtQQ("Excel 8.0;HDR=YES;IMEX=2;DATABASE=?", Fx)
    DrpTbl mTar(J), Db
    Db.TableDefs.Append Tbl
Next
End Sub

Function LnkTblFx$(T, Optional D As Database)
LnkTblFx = TakBet(Tbl(T, D).Connect, "Database=", ";")
End Function

Private Sub FxDb__Tst()
Dim Db As Database: Set Db = FxDb("N:\SapAccessReports\DutyPrepay5\SAPDownloadExcel\KE24 2010-01c.xls")
AyDmp DbTny(Db)
Db.Close
End Sub

Sub Tst()
FxDb__Tst
End Sub
