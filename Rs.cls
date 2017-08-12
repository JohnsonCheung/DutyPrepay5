VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Rs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Public Rs As Dao.Recordset
Function Init(A As Dao.Recordset) As Rs
Set Rs = A
Set Init = Me
End Function
Function Drs() As Drs
Set Drs = ccNew.Drs(Fny, Dry)
End Function

Function Dry() As Dry
Dim O As New Dry
With Rs
    While Not .EOF
        O.Push Flds.Dr
        .MoveNext
    Wend
End With
Set Dry = O
End Function

Function Fny() As String()
Fny = Flds.Fny
End Function

Function Sy() As String()
Dim O$()
With Rs
    While Not .EOF
        Push O$, Rs.Fields(0).Value
        .MoveNext
    Wend
End With
Sy = O
End Function

Function Flds() As Flds
Set Flds = ccNew.Flds(Rs.Fields)
End Function

