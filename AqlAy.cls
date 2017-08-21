VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AqlAy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Private SqlAy_$()
Public Cn_ As ADODB.Connection
Property Get SqlAy() As String()
SqlAy = SqlAy_
End Property
Function Init(SqlAy$(), Cn As ADODB.Connection) As AqlAy
SqlAy_ = SqlAy
Set Cn_ = Cn
Set Init = Me
End Function
Sub Run()
Dim I
For Each I In SqlAy
    Cn_.Execute I
Next
End Sub

