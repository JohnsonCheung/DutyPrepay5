Attribute VB_Name = "Xls_Wb"
Option Explicit
Option Compare Database

Function FxOpn(Fx) As Workbook
Set FxOpn = Xls.Workbooks.Open(Fx)
End Function

Function WbAddWs(a As Workbook, WsNm$) As Worksheet
Dim O As Worksheet
Set O = a.Sheets.Add(, WbLasWs(a))
O.Name = WsNm
Set WbAddWs = O
End Function

Sub WbClsNoSav(a As Workbook)
On Error Resume Next
a.Close False
End Sub

Function WbFstWs(a As Workbook) As Worksheet
Set WbFstWs = a.Sheets(1)
End Function

Function WbLasWs(a As Workbook) As Worksheet
Set WbLasWs = a.Sheets(a.Sheets.Count)
End Function

Function WbNew(Optional Vis As Boolean) As Workbook
Dim O As Workbook
Set O = Xls.Workbooks.Add
If Vis Then O.Visible = True
Set WbNew = O
End Function

Sub WbSav(a As Workbook)
Dim X As Excel.Application
Set X = a.Application
Dim Y As Boolean
Y = X.DisplayAlerts
X.DisplayAlerts = False
a.Save
X.DisplayAlerts = Y
End Sub

Sub WbVis(a As Workbook)
a.Application.Visible = True
End Sub
