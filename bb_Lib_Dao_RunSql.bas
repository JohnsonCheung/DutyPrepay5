Attribute VB_Name = "bb_Lib_Dao_RunSql"
Option Compare Database
Option Explicit

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
