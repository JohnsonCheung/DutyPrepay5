VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ws"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Public Ws As Worksheet
Function Init(Ws_ As Worksheet) As Ws
Set Ws = Ws_
Set Init = Me
End Function
Function Wb() As Workbook
Set Wb = Ws.Parent
End Function
Function A1X() As Rg
Set A1X = Rg(A1)
End Function
Function A1() As Range
Set A1 = Ws.Range("A1")
End Function
Function WbX() As Wb
Set WbX = Nw.Wb(Wb)
End Function
Sub ClsNoSav()
WbX.ClsNoSav
End Sub

Function CRR(C, R1, R2) As Range
Set CRR = RCRC(R1, C, R2, C)
End Function

Function DtaRg() As Range
Dim R, C
With LasCell
    R = .Row
    C = .Column
End With
Set DtaRg = RCRC(1, 1, R, C)
End Function
Function CrtLo(Optional LoNm$) As ListObject
Dim O As ListObject: Set O = Ws.ListObjects.Add(xlSrcRange, DtaRg, , xlYes)
If LoNm <> "" Then O.Name = LoNm
Lo(O).AdjColWdt
Set CrtLo = O
End Function
Function LasCell() As Range
Set LasCell = Ws.Cells.SpecialCells(xlCellTypeLastCell)
End Function

Function LasCno%()
LasCno = LasCell.Column
End Function

Function LasRno%()
LasRno = LasCell.Row
End Function

Function RC(R, C) As Range
Set RC = Ws.Cells(R, C)
End Function
Sub Dlt()
Ws.Delete
End Sub
Function RCRC(R1, C1, R2, C2) As Range
Set RCRC = Ws.Range(Ws.Cells(R1, C1), Ws.Cells(R2, C2))
End Function

Function Sq() As Variant()
Sq = DtaRg.Value
End Function

Sub Vis()
Ws.Application.Visible = True
End Sub
