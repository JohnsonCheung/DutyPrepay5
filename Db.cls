VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Db"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Public Db As Database
Function Init(Optional A As Database) As Db
Set Db = DftDb(A)
Set Init = Me
End Function
Function Ds(TblNmLvs_or_Tny, Optional DsNm$ = "Ds") As Ds
Dim Tny$(): Tny = Ny(TblNmLvs_or_Tny).Ny
Dim O As New Ds
    O.DsNm = DsNm
    Dim J%
    Dim U%: U = UB(Tny)
    For J = 0 To UB(Tny)
        O.AddDt Dt(Tny(J))
    Next
Set Ds = O
End Function

Sub BrwInf(Optional MaxColWdt% = 2000)
InfDs.Brw MaxColWdt, DsBrkLinMapStr:="TblFld:Tbl"
End Sub

Function Dt(T) As Dt
Set Dt = DbT(T).Dt
End Function
Sub Brw()
Dim N$: N = Db.Name
Db.Close
FbBrw (N)
End Sub

Function TblFldInfDt() As Dt
Dim Dry As New Dry
    Dim T, InfFny$()
    For Each T In Tny
        With DbT(T)
            If AyIsEmpty(InfFny) Then InfFny = .FldInfFny
            Dry.AddDr .FldInfDry.Dry
        End With
    Next
Set TblFldInfDt = Nw.Dt(InfFny, Dry, "TblFld")
End Function
Function DbT(T) As DbT
Set DbT = Nw.DbT(T, Db)
End Function
Function Tny() As String()
Const S$ = "Select Name from MSysObjects where Type in (1,6) and Left(Name,4)<>'MSYS'"
Tny = Sql(S).Sy
End Function
Function Sql(Sql_$, Optional Db As Database) As Sql
Set Sql = Nw.Sql(Sql_, Db)
End Function
Function StruInfDt() As Dt
Dim Dry As New Dry
    Dim T
    For Each T In Tny
        With DbT(T)
             Dry.AddDr Array(T, .RecCnt, .Des, .StruLin(SkipTblNm:=True))
        End With
    Next
Dim Fny$(): Fny = SplitSpc("Tbl RecCnt Des Stru")
Set StruInfDt = Nw.Dt(Fny, Dry, "Tbl")
End Function
Private Sub Qny__Tst()
Ay(Qny).Dmp
End Sub
Function Qny() As String()
Const S$ = "Select Name from MSysObjects where Type=5 and Left(Name,4)<>'MSYS' and Left(Name,4)<>'~sq_'"
Qny = Sql(S).Sy
End Function

Sub DrpTbl(T)
If IsTbl(T) Then Db.Execute FmtQQ("Drop Table [?]", T)
End Sub
Function Nm$()
Nm = Db.Name
End Function
Function InfDs() As Ds
Dim O As New Ds
O.AddDt LnkInfDt()
O.AddDt StruInfDt()
O.AddDt TblFldInfDt()
O.DsNm = Nm
Set InfDs = O
End Function
Function NTbl%()
Const S$ = "Select Count(*) from MSysObjects where Type in (1,6) and Left(Name,4)<>'MSYS'"""
NTbl = Sql(S).Lng
End Function
Function DbTy() As DbT()
Dim O() As DbT, J%, T
ReDim O(NTbl - 1)
For Each T In Tny
    Set O(J) = DbT(T)
    J = J + 1
Next
DbTy = O
End Function

Sub LnkFb(Fb, TblNm_or_Ny, Optional Tar)
Dim Src$(): Src = Ny(TblNm_or_Ny).Ny
Dim mTar$(): mTar = Ny(Dft(Tar, Src)).Ny
Dim Tbl As Dao.TableDef
Dim J%
For J = 0 To UB(mTar)
    Set Tbl = Db.CreateTableDef(mTar(J))
    Tbl.SourceTableName = Src(J)
    Tbl.Connect = FmtQQ(";DATABASE=?", Fb)
    DrpTbl mTar(J)
    Db.TableDefs.Append Tbl
Next
End Sub

Sub InsDs(A As Ds)
RunSqlAy InsDsSqlAy(A)
End Sub

Function InsDsSqlAy(A As Ds) As String()
If A.IsEmpty Then Exit Function
Dim O$()
Dim J%, DtAy() As Dt
DtAy = A.DtAy
For J = 0 To A.NDt - 1
    PushAy O, InsDtSqlAy(DtAy(J))
Next
InsDsSqlAy = O
End Function

Sub InsDt(A As Dt)
RunSqlAy InsDtSqlAy(A)
End Sub


Sub LnkFx(Fx, WsNm_or_Ny, Optional Tar)
Dim Src$(): Src = Ny(WsNm_or_Ny).Ny
Dim mTar$(): mTar = Ny(Dft(Tar, Src)).Ny
Dim Tbl As Dao.TableDef
Dim J%
For J = 0 To UB(mTar)
    Set Tbl = Db.CreateTableDef(mTar(J))
    Tbl.SourceTableName = Src(J) & "$"
    Tbl.Connect = FmtQQ("Excel 8.0;HDR=YES;IMEX=2;DATABASE=?", Fx)
    DrpTbl mTar(J)
    Db.TableDefs.Append Tbl
Next
End Sub
Sub CrtTbl(T, FldDclAy)
Db.Execute FmtQQ("Create Table [?] (?)", T, JnComma(FldDclAy))
End Sub

Sub RunSql(Sql)
Db.Execute Sql
End Sub
Function Fny(T) As String()
Fny = DbT(T).Fny
End Function

Sub RunSqlAy(SqlAy$())
If AyIsEmpty(SqlAy) Then Exit Sub
Dim Sql
For Each Sql In SqlAy
    RunSql Sql
Next
End Sub

Function InsDtSqlAy(Dt As Dt) As String()
Dim M As New Db_InsDt_SqlAy
InsDtSqlAy = M.SqlAy(Me, Dt)
End Function
Function DaoFld(T, F) As Dao.Field
Set DaoFld = DaoTbl(T).Fields(F)
End Function
Function DaoTbl(T) As Dao.TableDef
Set DaoTbl = Db.TableDefs(T)
End Function
Function LnkInfDt() As Dt
Dim Dry As New Dry
    Dim T, C$
    For Each T In Tny
        C = DaoTbl(T).Connect
        If C <> "" Then Dry.AddDr Array(T, C)
    Next
Dim Fny$(): Fny = ApSy("Tbl", "Connect")
Set LnkInfDt = Nw.Dt(Fny, Dry, "Lnk")
End Function

