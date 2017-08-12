Attribute VB_Name = "bb_UpgradeDb"
Option Compare Database
Option Explicit

Sub TblPermit_AddFld_IsImport()
Const T$ = "Permit"
Const F$ = "IsImport"
Dim Db As Database: Set Db = DtaDb
With DbT(T, Db)
    If .HasFld(F) Then Db.Close: Exit Sub
    .AddFld F, dbBoolean
End With
Log "Field [Permit.IsImport] is added"
Db.Close
End Sub
