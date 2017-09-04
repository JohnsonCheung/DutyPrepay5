Attribute VB_Name = "Xls_Ws"
Option Explicit
Option Compare Database

Function IsWs(A As Workbook, Idx_or_WsNm) As Boolean
On Error GoTo X
Dim Ws As Worksheet: Set Ws = A.Sheets(Idx_or_WsNm)
IsWs = True
Exit Function
X:
End Function

Sub LoAdjColWdt(A As ListObject)
Dim C As Range: Set C = LoEntCol(A)
C.AutoFit
Dim EntC As Range, J%
For J = 1 To C.Columns.Count
    Set EntC = RgEntC(C, J)
    If EntC.ColumnWidth > 100 Then EntC.ColumnWidth = 100
Next
End Sub

Function LoC(A As ListObject, C, Optional InclTot As Boolean, Optional InclHdr As Boolean) As Range
Dim mC%, R1&, R2&
R1 = LoR1(A, InclHdr)
R2 = LoR2(A, InclTot)
mC = LoWsCno(A, C)
Set LoC = WsCRR(LoWs(A), mC, R1, R2)
End Function

Function LoCC(A As ListObject, C1, C2, Optional InclTot As Boolean, Optional InclHdr As Boolean) As Range
Dim R1&, R2&, mC1%, mC2%
R1 = LoR1(A, InclHdr)
R2 = LoR2(A, InclTot)
mC1 = LoWsCno(A, C1)
mC2 = LoWsCno(A, C2)
Set LoCC = WsRCRC(LoWs(A), R1, mC1, R2, mC2)
End Function

Function LoCrt(A As Worksheet, Optional LoNm$) As ListObject
Dim R As Range: Set R = WsDtaRg(A)
If IsNothing(R) Then Exit Function
Dim O As ListObject: Set O = A.ListObjects.Add(xlSrcRange, WsDtaRg(A), , xlYes)
If LoNm <> "" Then O.Name = LoNm
LoAdjColWdt O
Set LoCrt = O
End Function

Function LoEntCol(A As ListObject) As Range
Set LoEntCol = LoCC(A, 1, LoNCol(A)).EntireColumn
End Function

Function LoIsNoDta(A As ListObject) As Boolean
LoIsNoDta = IsNothing(A.DataBodyRange)
End Function

Function LoNCol%(A As ListObject)
LoNCol = A.ListColumns.Count
End Function

Function LoR1&(A As ListObject, Optional InclHdr As Boolean)
If LoIsNoDta(A) Then
    LoR1 = A.ListColumns(1).Range.Row + 1
    Exit Function
End If
LoR1 = A.DataBodyRange.Row - IIf(InclHdr, 1, 0)
End Function

Function LoR2&(A As ListObject, Optional InclTot As Boolean)
If LoIsNoDta(A) Then
    LoR2 = LoR1(A)
    Exit Function
End If
LoR2 = A.DataBodyRange.Row + IIf(InclTot, 1, 0)
End Function

Sub LoVis(A As ListObject)
A.Application.Visible = True
End Sub

Function LoWs(A As ListObject) As Worksheet
Set LoWs = A.Parent
End Function

Function LoWsCno%(A As ListObject, Idx_or_ColNm)
LoWsCno = A.ListColumns(Idx_or_ColNm).Range.Column
End Function

Function WbWs(A As Workbook, Idx_or_WsNm) As Worksheet
Set WbWs = A.Sheets(Idx_or_WsNm)
End Function

Function WsA1(A As Worksheet) As Range
Set WsA1 = A.Range("A1")
End Function

Sub WsClsNoSav(A As Worksheet)
WbClsNoSav WsWb(A)
End Sub

Function WsCRR(A As Worksheet, C, R1, R2) As Range
Set WsCRR = WsRCRC(A, R1, C, R2, C)
End Function

Sub WsDlt(A As Workbook, Idx_or_WsNm)
If IsWs(A, Idx_or_WsNm) Then WbWs(A, Idx_or_WsNm).Delete
End Sub

Function WsDtaRg(A As Worksheet) As Range
Dim R, C
With WsLasCell(A)
    R = .Row
    C = .Column
End With
If R = 1 And C = 1 Then Exit Function
Set WsDtaRg = WsRCRC(A, 1, 1, R, C)
End Function

Function WsLasCell(A As Worksheet) As Range
Set WsLasCell = A.Cells.SpecialCells(xlCellTypeLastCell)
End Function

Function WsLasCno%(A As Worksheet)
WsLasCno = WsLasCell(A).Column
End Function

Function WsLasRno%(A As Worksheet)
WsLasRno = WsLasCell(A).Row
End Function

Function WsNew(Optional WsNm$, Optional Vis As Boolean) As Worksheet
Dim Wb As Workbook
Set Wb = WbNew
WsDlt Wb, "Sheet2"
WsDlt Wb, "Sheet3"
If WsNm <> "" Then WbWs(Wb, "Sheet1").Name = WsNm
Set WsNew = WbWs(Wb, 1)
If Vis Then WbVis Wb
End Function

Function WsRC(A As Worksheet, R, C) As Range
Set WsRC = A.Cells(R, C)
End Function

Function WsRCRC(A As Worksheet, R1, C1, R2, C2) As Range
Set WsRCRC = A.Range(A.Cells(R1, C1), A.Cells(R2, C2))
End Function

Function WsSq(A As Worksheet) As Variant()
WsSq = WsDtaRg(A).Value
End Function

Sub WsVis(A As Worksheet)
A.Application.Visible = True
End Sub

Function WsWb(A As Worksheet) As Workbook
Set WsWb = A.Parent
End Function

Sub LoAdjColWdt__Tst()
Dim Ws As Worksheet: Set Ws = WsNew(Vis:=True)
Dim Sq(1 To 2, 1 To 2)
Sq(1, 1) = "A"
Sq(1, 2) = "B"
Sq(2, 1) = "123123"
Sq(2, 2) = String(1234, "A")
Ws.Range("A1:B2").Value = Sq
LoAdjColWdt LoCrt(Ws)
WsClsNoSav Ws
End Sub

Sub Tst()
LoAdjColWdt__Tst
End Sub
