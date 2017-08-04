Attribute VB_Name = "bb_Lib_Xls_Wb"
Option Compare Database
Option Explicit
Function OpnWb(Fx) As Workbook
Set OpnWb = Xls.Workbooks.Open(Fx)
End Function
Function NewWb(Optional Vis As Boolean) As Workbook
Dim O As Workbook
Set O = Xls.Workbooks.Add
If Vis Then O.Visible = True
Set NewWb = O
End Function
Function LasWs(A As Workbook) As Worksheet
Set LasWs = A.Sheets(A.Sheets.Count)
End Function
Sub ClsWbNoSav(A As Workbook)
On Error Resume Next
A.Close False
End Sub
Sub WbVis(A As Workbook)
A.Application.Visible = True
End Sub
Function AddWs(A As Workbook, WsNm$) As Worksheet
Dim O As Worksheet
Set O = A.Sheets.Add(, LasWs(A))
O.Name = WsNm
Set AddWs = O
End Function
