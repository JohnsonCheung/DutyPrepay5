Attribute VB_Name = "bb_Lib_Xls"
Option Compare Database
Option Explicit
Property Get Xls() As Excel.Application
Static X As Excel.Application
On Error GoTo XX
Dim A$: A = X.Name
Set Xls = X
Exit Property
XX:
Set X = New Excel.Application
Set Xls = X
End Property
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
Function ReSzRg(Cell As Range, Sq) As Range
Dim R, C
R = UBound(Sq, 1)
C = UBound(Sq, 2)
Set ReSzRg = RgRCRC(Cell, 1, 1, R, C)
End Function
