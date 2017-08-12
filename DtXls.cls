VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DtXls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Public Dt As Dt
Function Init(A) As DtXls
If TypeName(A) <> "Dt" Then Stop
Set Dt = A
End Function
Function AddWb(Wb As Workbook) As Worksheet
Dim O As Worksheet
Set O = WbAddWs(Wb, Dt.DtNm)
PutAt WsA1(O)
Set AddWb = O
End Function
Sub PutAt(At As Range)

End Sub

Function Ws(Optional Hid As Boolean) As Worksheet
Dim O As Worksheet
Set O = WsNew(Dt.DtNm, Not Hid)
PutAt WsA1(O)
Set Ws = O
End Function

