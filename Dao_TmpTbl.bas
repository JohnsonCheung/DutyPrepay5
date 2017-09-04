Attribute VB_Name = "Dao_TmpTbl"
Option Explicit
Option Compare Database

Sub EnsTbl_Tmp1()
If IsTbl("Tmp1") Then Exit Sub
RunSql "Create Table Tmp1 (AA Int, BB Text 10)"
End Sub
