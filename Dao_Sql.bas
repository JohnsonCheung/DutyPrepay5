Attribute VB_Name = "Dao_Sql"
Option Compare Database
Option Explicit

Function SqlBool(Sql$, Optional D As Database) As Boolean
SqlBool = DftDb(D).OpenRecordset(Sql).EOF
End Function

Function SqlDrs(Sql$, Optional D As Database) As Drs
Dim Rs As Recordset
Dim O As Drs
Set Rs = DftDb(D).OpenRecordset(Sql)
O.Dry = RsDry(Rs)
O.Fny = RsFny(Rs)
SqlDrs = O
End Function

Function SqlDry(Sql$, Optional D As Database) As Variant()
SqlDry = RsDry(DftDb(D).OpenRecordset(Sql))
End Function

Function SqlLng&(Sql$, Optional D As Database)
SqlLng = SqlV(Sql, D)
End Function

Function SqlRs(Sql$, Optional D As Database) As Dao.Recordset
Set SqlRs = DftDb(D).OpenRecordset(Sql)
End Function

Function SqlSy(Sql$, Optional D As Database) As String()
SqlSy = RsSy(DftDb(D).OpenRecordset(Sql))
End Function

Function SqlV(Sql$, Optional D As Database)
With SqlRs(Sql, D)
    SqlV = .Fields(0).Value
    .Close
End With
End Function
