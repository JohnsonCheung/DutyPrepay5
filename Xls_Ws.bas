Attribute VB_Name = "Xls_Ws"
Option Explicit
Option Compare Database

Function IsWs(a As Workbook, Idx_or_WsNm) As Boolean
On Error GoTo X
Dim Ws As Worksheet: Set Ws = a.Sheets(Idx_or_WsNm)
IsWs = True
Exit Function
X:
End Function

Sub LoAdjColWdt(a As ListObject)
Dim C As Range: Set C = LoEntCol(a)
C.AutoFit
Dim EntC As Range, J%
For J = 1 To C.Columns.Count
    Set EntC = RgEntC(C, J)
    If EntC.ColumnWidth > 100 Then EntC.ColumnWidth = 100
Next
End Sub

Function LoC(a As ListObject, C, Optional InclTot As Boolean, Optional InclHdr As Boolean) As Range
Dim mC%, R1&, R2&
R1 = LoR1(a, InclHdr)
R2 = LoR2(a, InclTot)
mC = LoWsCno(a, C)
Set LoC = WsCRR(LoWs(a), mC, R1, R2)
End Function

Function LoCC(a As ListObject, C1, C2, Optional InclTot As Boolean, Optional InclHdr As Boolean) As Range
Dim R1&, R2&, mC1%, mC2%
R1 = LoR1(a, InclHdr)
R2 = LoR2(a, InclTot)
mC1 = LoWsCno(a, C1)
mC2 = LoWsCno(a, C2)
Set LoCC = WsRCRC(LoWs(a), R1, mC1, R2, mC2)
End Function

Function LoCrt(a As Worksheet, Optional LoNm$) As ListObject
Dim R As Range: Set R = WsDtaRg(a)
If IsNothing(R) Then Exit Function
Dim O As ListObject: Set O = a.ListObjects.Add(xlSrcRange, WsDtaRg(a), , xlYes)
If LoNm <> "" Then O.Name = LoNm
LoAdjColWdt O
Set LoCrt = O
End Function

Function LoEntCol(a As ListObject) As Range
Set LoEntCol = LoCC(a, 1, LoNCol(a)).EntireColumn
End Function

Function LoIsNoDta(a As ListObject) As Boolean
LoIsNoDta = IsNothing(a.DataBodyRange)
End Function

Function LoNCol%(a As ListObject)
LoNCol = a.ListColumns.Count
End Function

Function LoR1&(a As ListObject, Optional InclHdr As Boolean)
If LoIsNoDta(a) Then
    LoR1 = a.ListColumns(1).Range.Row + 1
    Exit Function
End If
LoR1 = a.DataBodyRange.Row - IIf(InclHdr, 1, 0)
End Function

Function LoR2&(a As ListObject, Optional InclTot As Boolean)
If LoIsNoDta(a) Then
    LoR2 = LoR1(a)
    Exit Function
End If
LoR2 = a.DataBodyRange.Row + IIf(InclTot, 1, 0)
End Function

Sub LoVis(a As ListObject)
a.Application.Visible = True
End Sub

Function LoWs(a As ListObject) As Worksheet
Set LoWs = a.Parent
End Function

Function LoWsCno%(a As ListObject, Idx_or_ColNm)
LoWsCno = a.ListColumns(Idx_or_ColNm).Range.Column
End Function

Function WbWs(a As Workbook, Idx_or_WsNm) As Worksheet
Set WbWs = a.Sheets(Idx_or_WsNm)
End Function

Function WsA1(a As Worksheet) As Range
Set WsA1 = a.Range("A1")
End Function

Sub WsClsNoSav(a As Worksheet)
WbClsNoSav WsWb(a)
End Sub

Function WsCRR(a As Worksheet, C, R1, R2) As Range
Set WsCRR = WsRCRC(a, R1, C, R2, C)
End Function

Sub WsDlt(a As Workbook, Idx_or_WsNm)
If IsWs(a, Idx_or_WsNm) Then WbWs(a, Idx_or_WsNm).Delete
End Sub

Function WsDtaRg(a As Worksheet) As Range
Dim R, C
With WsLasCell(a)
    R = .Row
    C = .Column
End With
If R = 1 And C = 1 Then Exit Function
Set WsDtaRg = WsRCRC(a, 1, 1, R, C)
End Function

Function WsLasCell(a As Worksheet) As Range
Set WsLasCell = a.Cells.SpecialCells(xlCellTypeLastCell)
End Function

Function WsLasCno%(a As Worksheet)
WsLasCno = WsLasCell(a).Column
End Function

Function WsLasRno%(a As Worksheet)
WsLasRno = WsLasCell(a).Row
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

Function WsRC(a As Worksheet, R, C) As Range
Set WsRC = a.Cells(R, C)
End Function

Function WsRCRC(a As Worksheet, R1, C1, R2, C2) As Range
Set WsRCRC = a.Range(a.Cells(R1, C1), a.Cells(R2, C2))
End Function

Function WsSq(a As Worksheet) As Variant()
WsSq = WsDtaRg(a).Value
End Function

Sub WsVis(a As Worksheet)
a.Application.Visible = True
End Sub

Function WsWb(a As Worksheet) As Workbook
Set WsWb = a.Parent
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
