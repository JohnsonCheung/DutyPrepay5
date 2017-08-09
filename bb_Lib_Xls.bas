Attribute VB_Name = "bb_Lib_Xls"
Option Compare Database
Option Explicit

Function FxHasWs(Fx, WsNm) As Boolean
FxHasWs = AyHas(FxWsNy(Fx), WsNm)
End Function

Sub FxRmvWsIfExist(Fx, WsNm)
If FxHasWs(Fx, WsNm) Then
    Dim B As Workbook: Set B = FxOpn(Fx)
    WbWs(B, WsNm).Delete
    WbSav B
    WbClsNoSav B
End If
End Sub

Function ReSzRg(Cell As Range, Sq) As Range
Dim R, C
R = UBound(Sq, 1)
C = UBound(Sq, 2)
Set ReSzRg = RgRCRC(Cell, 1, 1, R, C)
End Function

Function RgC(A As Range, C) As Range
Set RgC = RgCRR(A, C, 1, A.Rows.Count)
End Function

Function RgCRR(A As Range, C, R1, R2) As Range
Set RgCRR = RgRCRC(A, R1, C, R2, C)
End Function

Function RgEntC(A As Range, C) As Range
Set RgEntC = RgC(A, C).EntireColumn
End Function

Function RgRC(Rg As Range, R, C) As Range
Set RgRC = Rg.Cells(R, C)
End Function

Function RgRCRC(Rg As Range, R1, C1, R2, C2) As Range
Dim Ws As Worksheet, Cell1 As Range, Cell2 As Range
Set Ws = Rg.Parent
Set Cell1 = RgRC(Rg, R1, C1)
Set Cell2 = RgRC(Rg, R2, C2)
Set RgRCRC = Ws.Range(Cell1, Cell2)
End Function

Function RgWs(A As Range) As Worksheet
Set RgWs = A.Parent
End Function

Function Xls() As Excel.Application
Static X As Excel.Application
On Error GoTo XX
Dim A$: A = X.Name
Set Xls = X
Exit Function
XX:
Set X = New Excel.Application
Set Xls = X
End Function

Sub FxRmvWsIfExist__Tst()
Dim T$: T = TmpFx
Dim Wb As Workbook
Set Wb = WbNew
Wb.Sheets.Add
Wb.SaveAs T
Dim WsNyBef$(), WsNyAft$()
    WsNyBef = FxWsNy(T)
    FxRmvWsIfExist T, "Sheet1"
    WsNyAft = FxWsNy(T)
Dim Exp$()
    Exp = AyMinus(WsNyBef, Array("Sheet1"))
AssertEqAy Exp, WsNyAft
End Sub

Sub Tst()
FxRmvWsIfExist__Tst
End Sub
