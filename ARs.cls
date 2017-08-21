VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ARs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Public Rs_ As ADODB.Recordset
Function Init(Rs As ADODB.Recordset) As ARs
Set Rs_ = Rs
Set Init = Me
End Function
Function Drs() As Drs
Set Drs = Nw.Drs(Fny, Dry)
End Function

Function Dry() As Dry
Dim O As New Dry
With Rs_
    While Not .EOF
        O.AddDr AFlds.Dr
        .MoveNext
    Wend
End With
Set Dry = O
End Function

Function Fny() As String()
Fny = AFlds.Fny
End Function

Function AFlds() As AFlds
Set AFlds = Nw.AFlds(Rs_.Fields)
End Function

