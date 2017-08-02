Attribute VB_Name = "bb_Lib_Xls_Ws"
Option Compare Database
Option Explicit
Function WsSq(A As Worksheet) As Variant()
WsSq = WsDtaRg(A).Value
End Function
Function WsLasCell(A As Worksheet) As Range
Set WsLasCell = A.Cells.SpecialCells(xlCellTypeLastCell)
End Function
Function WsDtaRg(A As Worksheet) As Range
Dim R, C
With WsLasCell(A)
    R = .Row
    C = .Column
End With
Set WsDtaRg = WsRCRC(A, 1, 1, R, C)
End Function
Sub ClsWsNoSav(A As Worksheet)
ClsWbNoSav WsWb(A)
End Sub
Sub DltWs(A As Workbook, Idx_or_WsNm)
If IsWs(A, Idx_or_WsNm) Then WbWs(A, Idx_or_WsNm).Delete
End Sub
Function WbWs(A As Workbook, Idx_or_WsNm) As Worksheet
Set WbWs = A.Sheets(Idx_or_WsNm)
End Function
Function IsWs(A As Workbook, Idx_or_WsNm) As Boolean
On Error GoTo X
Dim Ws As Worksheet: Set Ws = A.Sheets(Idx_or_WsNm)
IsWs = True
Exit Function
X:
End Function
Function NewWs(Optional WsNm$, Optional Vis As Boolean) As Worksheet
Dim Wb As Workbook
Set Wb = NewWb
DltWs Wb, "Sheet2"
DltWs Wb, "Sheet3"
If WsNm <> "" Then WbWs(Wb, "Sheet1").Name = WsNm
Set NewWs = WbWs(Wb, 1)
If Vis Then WbVis Wb
End Function
Function WsRCRC(A As Worksheet, R1, C1, R2, C2) As Range
Set WsRCRC = A.Range(A.Cells(R1, C1), A.Cells(R2, C2))
End Function
Function WsWb(A As Worksheet) As Workbook
Set WsWb = A.Parent
End Function
Function WsA1(A As Worksheet) As Range
Set WsA1 = A.Range("A1")
End Function
Function WsRC(A As Worksheet, R, C) As Range
Set WsRC = A.Cells(R, C)
End Function
Sub WsVis(A As Worksheet)
A.Application.Visible = True
End Sub
