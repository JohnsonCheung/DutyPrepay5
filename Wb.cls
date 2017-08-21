VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Wb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Public Wb_ As Workbook
Function Init(A As Workbook) As Wb
Set Wb_ = A
Set Init = Me
End Function
Function AddWsX(WsNm$) As Ws
Set AddWsX = Ws(AddWs(WsNm))
End Function
Function Ws(Idx_or_WsNm) As Worksheet
Set Ws = Wb_.Sheets(Idx_or_WsNm)
End Function
Function IsWs(Idx_or_WsNm) As Boolean
On Error GoTo X
Dim Ws As Worksheet: Set Ws = Wb_.Sheets(Idx_or_WsNm)
IsWs = True
Exit Function
X:
End Function

Function WsX(Idx_or_WsNm) As Ws
Set WsX = Nw.Ws(Ws(Idx_or_WsNm))
End Function

Function AddWs(WsNm$) As Worksheet
Dim O As Worksheet
Set O = Wb_.Sheets.Add(, LasWs)
O.Name = WsNm
Set AddWs = O
End Function
Sub Tst()

End Sub
Sub ClsNoSav()
On Error Resume Next
Wb_.Close False
End Sub
Function DltWs(Idx_or_WsNm) As Wb
If IsWs(Idx_or_WsNm) Then WsX(Idx_or_WsNm).Dlt
Set DltWs = Me
End Function

Function FstWs() As Worksheet
Set FstWs = Wb_.Sheets(1)
End Function

Function LasWs() As Worksheet
Set LasWs = Wb_.Sheets(Wb_.Sheets.Count)
End Function

Function Sav() As Wb
Dim X As Excel.Application
Set X = Wb_.Application
Dim Y As Boolean
Y = X.DisplayAlerts
X.DisplayAlerts = False
Wb_.Save
X.DisplayAlerts = Y
Set Sav = Me
End Function

Sub Vis()
Wb_.Application.Visible = True
End Sub

