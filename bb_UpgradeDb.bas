Attribute VB_Name = "bb_UpgradeDb"
Option Compare Database
Option Explicit

Sub TblPermit_AddFld_IsImport()
Const T$ = "Permit"
Const F$ = "IsImport"
Dim Db As Database: Set Db = DtaDb
If HasFld(T, F, Db) Then Db.Close: Exit Sub
AddFld T, F, dbBoolean, Db
Log "Field [Permit.IsImport] is added"
Db.Close
End Sub
