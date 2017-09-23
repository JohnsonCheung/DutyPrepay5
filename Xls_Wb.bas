Attribute VB_Name = "Xls_Wb"
Option Explicit
Option Compare Database

Function FxOpn(Fx) As Workbook
Set FxOpn = Xls.Workbooks.Open(Fx)
End Function

Function WbAddWs(A As Workbook, Optional WsNm$) As Worksheet
Dim O As Worksheet
Set O = A.Sheets.Add(, WbLasWs(A))
If WsNm <> "" Then
    O.Name = WsNm
End If
Set WbAddWs = O
End Function

Sub WbClsNoSav(A As Workbook)
On Error Resume Next
A.Close False
End Sub

Function WbFstWs(A As Workbook) As Worksheet
Set WbFstWs = A.Sheets(1)
End Function

Function WbLasWs(A As Workbook) As Worksheet
Set WbLasWs = A.Sheets(A.Sheets.Count)
End Function

Function WbNew(Optional Vis As Boolean) As Workbook
Dim O As Workbook
Set O = Xls.Workbooks.Add
If Vis Then O.Visible = True
Set WbNew = O
End Function

Sub WbSav(A As Workbook)
Dim X As Excel.Application
Set X = A.Application
Dim Y As Boolean
Y = X.DisplayAlerts
X.DisplayAlerts = False
A.Save
X.DisplayAlerts = Y
End Sub

Sub WbVis(A As Workbook)
A.Application.Visible = True
End Sub
