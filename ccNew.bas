Attribute VB_Name = "ccNew"
Option Compare Database
Function Rel() As Rel
Set Rel = New Rel
End Function

Function Md(A As CodeModule) As Md
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
Function PjByNm(Optional PjNm$) As Pj
Set PjByNm = Pj(Application.VBE.VBProjects(DftPjNm))
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
Function SqXls(Sq As Sq) As SqXls
Dim O As New SqXls
Set SqXls = O.Init(Sq)
End Function
Function DbT(TblNm, Optional Db As Database) As DbT
Dim O As New DbT
Set DbT = O.Init(TblNm, Db)
End Function
Function DsXls(A As Ds) As DsXls
Dim O As New DsXls
Set DsXls = O.Init(A)
End Function
Function DtXls(A As Dt) As DtXls
Dim O As New DtXls
Set DtXls = O.Init(A)
End Function
Function DrsXls(A As Drs) As DrsXls
Dim O As New DrsXls
Set DrsXls = O.Init(A)
End Function
Function ARs(A As ADODB.Recordset) As ARs
Dim O As New ARs
Set ARs = O.Init(A)
End Function
Function SqByHAy(HAy) As Sq
Dim O As New Sq
Set SqByHAy = O.InitByHAy(HAy)
End Function
Function SrcLin(L) As SrcLin
Dim O As New SrcLin
Set SrcLin = O.Init(L)
End Function
Function Dry() As Dry
Set Dry = New Dry
End Function
Function DftDry(A As Dry) As Dry
If IsNothing(A) Then
    Set DftDry = New Dry
Else
    Set DftDry = A
End If
End Function
Function Sy(Sy_$()) As Sy
Dim O As New Sy
Set Sy = O.Init(Sy_)
End Function
Function Ay(Ay_) As Ay
Dim O As New Ay
Set Ay = O.Init(Ay_)
End Function
Function Drs(Fny$(), Optional Dry As Dry) As Drs
Dim O As New Drs
Set Drs = O.Init(Fny, Dry)
End Function
