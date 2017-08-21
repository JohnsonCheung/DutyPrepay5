VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Rg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Public Rg As Range
Function Init(A As Range) As Rg
Set Rg = A
Set Init = Me
End Function
Function Ws() As Worksheet
Set Ws = Rg.Parent
End Function
Function WsX() As Ws
Set WsX = Nw.Ws(Ws)
End Function
Function Row(R) As Range
Set Row = RCC(R, 1, Rg.Columns.Count)
End Function
Function Col(C) As Range
Set Col = CRR(C, 1, Rg.Rows.Count)
End Function

Function CRR(C, R1, R2) As Range
Set CRR = RCRC(R1, C, R2, C)
End Function

Function RCC(R, C1, C2) As Range
Set RCC = RCRC(R, C1, R, C2)
End Function

Function EntC(C) As Range
Set EntC = Col(C).EntireColumn
End Function

Function RC(R, C) As Range
Set RC = Rg.Cells(R, C)
End Function

Function ReSz(Sq) As Range
Dim R, C
R = UBound(Sq, 1)
C = UBound(Sq, 2)
Set ReSz = RCRC(1, 1, R, C)
End Function

Function RCRC(R1, C1, R2, C2) As Range
Dim Cell1 As Range, Cell2 As Range
Set Cell1 = RC(R1, C1)
Set Cell2 = RC(R2, C2)
Set RCRC = Ws.Range(Cell1, Cell2)
End Function


