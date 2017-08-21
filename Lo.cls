VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Lo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Public Lo As ListObject
Function AdjColWdt() As ListObject
Dim C As Range: Set C = AllCol
C.AutoFit
Dim EntC As Range, J%
For J = 1 To C.Columns.Count
    Set EntC = Rg(C).EntC(J)
    If EntC.ColumnWidth > 100 Then EntC.ColumnWidth = 100
Next
Set AdjColWdt = Me
End Function
Function Col(C, Optional InclTot As Boolean, Optional InclHdr As Boolean) As Range
Dim mC%, mR1&, mR2&
mR1 = R1(InclHdr)
mR2 = R2(InclTot)
mC = Cno(C)
Set Col = WsX.CRR(mC, R1, R2)
End Function
Function Cno%(Idx_or_ColNm)
Cno = Lo.ListColumns(Idx_or_ColNm).Range.Column
End Function
Function WsX() As Ws
Set WsX = Nw.Ws(Ws)
End Function
Function CC(C1, C2, Optional InclTot As Boolean, Optional InclHdr As Boolean) As Range
Dim mR1&, mR2&, mC1%, mC2%
mR1 = R1(InclHdr)
mR2 = R2(InclTot)
mC1 = Cno(C1)
mC2 = Cno(C2)
Set CC = WsX.RCRC(mR1, mC1, mR2, mC2)
End Function
Function Init(Lo_ As ListObject) As Lo
Set Lo = Lo_
Set Init = Me
End Function
Sub AdjColWdt__Tst()
Dim Ws As Worksheet: Set Ws = NewWs(Vis:=True)
Dim Sq(1 To 2, 1 To 2)
Sq(1, 1) = "A"
Sq(1, 2) = "B"
Sq(2, 1) = "123123"
Sq(2, 2) = String(1234, "A")
Ws.Range("A1:B2").Value = Sq
ZZ.AdjColWdt
'WsClsNoSav Ws
End Sub
Function ZZ() As Lo

End Function
Function AllCol(Optional InclTot As Boolean, Optional InclHdr As Boolean) As Range
Set AllCol = CC(1, NCol, InclTot, InclHdr)
End Function

Function NCol%()
NCol = Lo.ListColumns.Count
End Function
Function R1&(Optional InclHdr As Boolean)
R1 = Lo.DataBodyRange.Row - IIf(InclHdr, 1, 0)
End Function

Function R2&(Optional InclTot As Boolean)
R2 = Lo.DataBodyRange.Row + IIf(InclTot, 1, 0)
End Function

Sub Vis()
Lo.Application.Visible = True
End Sub

Function Ws() As Worksheet
Set Ws = Lo.Parent
End Function




