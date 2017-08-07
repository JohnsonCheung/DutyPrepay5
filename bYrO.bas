Attribute VB_Name = "bYrO"
Option Compare Database
Option Base 0
Option Explicit

Sub AddYrO()
'Aim: There is no current Yr record in table YrO, create one record in YrO
With CurrentDb.OpenRecordset("Select Yr from YrO where Yr=" & VBA.Year(Date) - 2000)
    If Not .EOF Then .Close: Exit Sub
    .Close
End With
DoCmd.RunSql "Insert Into YrO (Yr) values (Year(Date())-2000)"
End Sub

Public Function IsLasYrOD_Exist(pY As Byte) As Boolean
Dim mYr As Byte
mYr = pY - 1
With CurrentDb.OpenRecordset("Select count(*) from YrOD where Yr=" & mYr)
    If Nz(.Fields(0).Value, 0) > 0 Then IsLasYrOD_Exist = True
End With
End Function
