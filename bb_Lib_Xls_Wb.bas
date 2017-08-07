Attribute VB_Name = "bb_Lib_Xls_Wb"
Option Compare Database
Option Explicit

Function LasWs(A As Workbook) As Worksheet
Set LasWs = A.Sheets(A.Sheets.Count)
End Function

Function WbAddWs(A As Workbook, WsNm$) As Worksheet
Dim O As Worksheet
Set O = A.Sheets.Add(, LasWs(A))
O.Name = WsNm
Set WbAddWs = O
End Function

Sub WbClsNoSav(A As Workbook)
On Error Resume Next
A.Close False
End Sub

Function WbNew(Optional Vis As Boolean) As Workbook
Dim O As Workbook
Set O = Xls.Workbooks.Add
If Vis Then O.Visible = True
Set WbNew = O
End Function

Function WbOpn(Fx) As Workbook
Set WbOpn = Xls.Workbooks.Open(Fx)
End Function

Sub WbVis(A As Workbook)
A.Application.Visible = True
End Sub
