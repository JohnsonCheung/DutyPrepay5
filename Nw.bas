Attribute VB_Name = "Nw"
Option Compare Database
Option Explicit
Function Rel() As Rel
Set Rel = New Rel
End Function
Function Md(Optional A As CodeModule) As Md
Dim O As New Md
Set Md = O.Init(A)
End Function
Function MdByNm(MdNm$, Optional PjNm$) As Md
Set MdByNm = PjByNm(PjNm).Md(MdNm)
End Function

Function Pj(Optional A As VBProject) As Pj
Dim O As New Pj
Set Pj = O.Init(A)
End Function
Function PjByNm(PjNm$) As Pj
Set PjByNm = Pj(Application.VBE.VBProjects(PjNm))
End Function
Function Lo(Lo_ As ListObject) As Lo
Dim O As New Lo
Set Lo = O.Init(Lo_)
End Function
Function Dt(Fny$(), Dry As Dry, Optional DtNm = "Dt") As Dt
Dim O As New Dt
Set Dt = O.Init(Fny, Dry, DtNm)
End Function
Function Sql3(Ly$()) As Sql3
Dim O As New Sql3
Set Sql3 = O.Init(Ly)
End Function
Function Rs(A As Dao.Recordset) As Rs
Dim O As New Rs
Set Rs = O.Init(A)
End Function
Function Fld(A As Dao.Field) As Fld
Dim O As New Fld
Set Fld = O.Init(A)
End Function
Function Sq(Sq_()) As Sq
Dim O As New Sq
Set Sq = O.Init(Sq_)
End Function
Function Db(Optional A As Database) As Db
Dim O As New Db
Set Db = O.Init(A)
End Function

Function Prps(A As Dao.Properties) As Prps
Dim O As New Prps
Set Prps = O.Init(A)
End Function
Function Flds(A As Dao.Fields) As Flds
Dim O As New Flds
Set Flds = O.Init(A)
End Function
Function Tbl(A As Dao.TableDef) As Tbl
Dim O As New Tbl
Set Tbl = O.Init(A)
End Function
Function Sql(Sql_$, D As Database) As Sql
Dim O As New Sql
Set Sql = O.Init(Sql_, D)
End Function
Function Ny(Optional Ny_or_NmLvs) As Ny
Dim O As New Ny
Set Ny = O.Init(Ny_or_NmLvs)
End Function
Function DbT(TblNm, Optional Db As Database) As DbT
Dim O As New DbT
Set DbT = O.Init(TblNm, Db)
End Function
Function Rg(A As Range) As Rg
Dim O As New Rg
Set Rg = O.Init(A)
End Function
Function Dic(Optional Dic_ As Dictionary) As Dic
Dim O As New Dic
Set Dic = O.Init(Dic_)
End Function
Function DicByMapStr(MapStr$) As Dic
Dim O As New Dic
'Set Dic = O.InitByMapStr(MapStr$)
End Function
Function ARs(A As ADODB.Recordset) As ARs
Dim O As New ARs
Set ARs = O.Init(A)
End Function
Function Pth(P) As Pth
Dim O As New Pth
Set Pth = O.Init(P)
End Function
Function BrkLinByDrsLy(DrsLy$(), BrkColNm$) As BrkLin
Dim O As New BrkLin
Set BrkLinByDrsLy = O.InitByDrsLy(DrsLy, BrkColNm)
End Function
Function BrkLinByDryLy(DryLy$(), BrkColIdx%) As BrkLin
Dim O As New BrkLin
Set BrkLinByDryLy = O.InitByDryLy(DryLy, BrkColIdx)
End Function

Function Map(Sy1, Sy2) As Map
Dim O As New Map
Set Map = O.Init(Sy1, Sy2)
End Function
Function SqByHAy(HAy) As Sq
Dim O As New Sq
Set SqByHAy = O.InitByHAy(HAy)
End Function
Function Ws(Ws_ As Worksheet) As Ws
Dim O As New Ws
Set Ws = O.Init(Ws_)
End Function
Function Ffn(Ffn_) As Ffn
Dim O As New Ffn
Set Ffn = O.Init(Ffn_)
End Function
Function Fx(Fx_) As Fx
Dim O As New Fx
Set Fx = O.Init(Fx_)
End Function
Function StrOpt(Optional S) As StrOpt
Dim O As New StrOpt
Set StrOpt = O.Init(S)
End Function
Function BoolOpt(Optional Bool) As BoolOpt
Dim O As New BoolOpt
Set BoolOpt = O.Init(Bool)
End Function
Function Aql(Sql, Optional Cn As ADODB.Connection) As Aql
Dim O As New Aql
Set Aql = O.Init(Sql, Cn)
End Function
Function SyOpt(Optional Sy) As SyOpt
Dim O As New SyOpt
Set SyOpt = O.Init(Sy)
End Function
Function AqlAy(SqlAy$(), Cn As ADODB.Connection) As ADODB.Connection
Dim O As New AqlAy
Set AqlAy = O.Init(SqlAy, Cn)
End Function
Function ACat(Cat As Catalog) As ACat
Dim O As New ACat
Set ACat = O.Init(Cat)
End Function
Function OAy() As Ay
Set OAy = New Ay
End Function
Function SqByVAy(VAy) As Sq
Dim O As New Sq
Set SqByVAy = O.InitByVAy(VAy)
End Function
Function PermImpFx(Fx) As PermImpFx
Dim O As New PermImpFx
Set PermImpFx = O.Init(Fx)
End Function
Function OWb() As Wb
Set OWb = New Wb
End Function
Function Wb(A As Workbook) As Wb
Set Wb = OWb.Init(A)
End Function
Function AFlds(A As ADODB.Fields) As AFlds
Dim O As New AFlds
Set AFlds = O.Init(A)
End Function
Function Parser(L) As Parser
Dim O As New Parser
Set Parser = O.Init(L)
End Function
Function Ft(Ft_) As Ft
Dim O As New Ft
Set Ft = O.Init(Ft_)
End Function
Function OSrcLin() As SrcLin
Set OSrcLin = New SrcLin
End Function
Function SrcLin(L) As SrcLin
Set SrcLin = OSrcLin.Init(L)
End Function
Function ODry() As Dry
Set ODry = New Dry
End Function
Function Dry(Dry_) As Dry
Set Dry = ODry.Init(Dry_)
End Function
Function Ay(Ay_) As Ay
Set Ay = OAy.Init(Ay_)
End Function
Function Drs(Fny$(), Optional Dry As Dry) As Drs
Dim O As New Drs
Set Drs = O.Init(Fny, Dry)
End Function


