Attribute VB_Name = "Dao_RunSql"
Option Explicit
Option Compare Database

Sub RunSql(Sql, Optional D As Database)
DftDb(D).Execute Sql
End Sub

Sub RunSqlAy(SqlAy$(), Optional D As Database)
If AyIsEmpty(SqlAy) Then Exit Sub
Dim Db As Database: Set Db = DftDb(D)
Dim Sql
For Each Sql In SqlAy
    RunSql Sql, D
Next
End Sub
