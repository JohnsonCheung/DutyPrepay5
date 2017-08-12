VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sql"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Private D As Database
Private A$
Function Init(Sql$, Optional Db As Database) As Sql
Set D = DftDb(D)
A = Sql
Set Init = Me
End Function
Sub Brw(Optional MaxColWdt% = 100)
Drs.Brw MaxColWdt
End Sub
Function Drs() As Drs
With Rs
Set Drs = ccNew.Drs(.Fny, .Dry)
End With
End Function
Function Rs() As Rs
Set Rs = ccNew.Rs(DaoRs)
End Function
Function DaoRs() As Dao.Recordset
On Error Resume Next
Set DaoRs = D.OpenRecordset(A)
End Function
Function Dry() As Dry
Set Dry = Rs.Dry
End Function

Function Lng&()
Lng = V
End Function

Function Sy() As String()
Sy = Rs.Sy
End Function

Function V()
On Error Resume Next
With Rs.Rs
    V = .Fields(0).Value
    .Close
End With
End Function

