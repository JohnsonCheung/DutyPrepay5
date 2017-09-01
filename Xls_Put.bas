Attribute VB_Name = "Xls_Put"
Option Compare Database
Option Explicit

Sub AyPut(Ay, Cell As Range)
SqPut Cell, AySqH(Ay)
End Sub

Function AySqH(Ay) As Variant()
If AyIsEmpty(Ay) Then Exit Function
Dim O(), C%
ReDim O(1 To 1, 1 To Sz(Ay))
C = 0
Dim V
For Each V In Ay
    C = C + 1
    O(1, C) = V
Next
AySqH = O
End Function

Sub DrsPut(A As Drs, At As Range, Optional LoNm$)
AyPut A.Fny, At
SqPut RgRC(At, 2, 1), DrySq(A.Dry, Sz(A.Fny))
LoCrt RgWs(At), LoNm
End Sub

Function DrsWs(A As Drs, Optional WsNm$ = "Sheet1") As Worksheet
Dim O As Worksheet: Set O = WsNew(WsNm, Vis:=True)
DrsPut A, WsA1(O)
Set DrsWs = O
End Function

Sub DryPut(AtCell As Range, Dry)
AtCell.Value = DrySq(Dry)
End Sub

Function DsNDt%(A As Ds)
DsNDt = DtAySz(A.DtAy)
End Function

Function DsWb(A As Ds) As Workbook
Dim O As Workbook
Set O = WbNew
With WbFstWs(O)
    .Name = "Ds"
    .Range("A1").Value = A.DsNm
End With
If Not DsIsEmpty(A) Then
    Dim J%
    For J = 0 To DsNDt(A) - 1
        WbAddDt O, A.DtAy(J)
    Next
End If
Set DsWb = O
End Function

Function DtWs(A As Dt) As Worksheet
Dim O As Worksheet
Set O = WsNew(A.DtNm)
DrsPut DtDrs(A), WsA1(O)
Set DtWs = O
End Function

Sub SqPut(Cell As Range, Sq)
If SqIsEmpty(Sq) Then Exit Sub
ReSzRg(Cell, Sq).Value = Sq
End Sub

Function TblWs(T, Optional D As Database) As Worksheet
Set TblWs = DtWs(TblDt(T, D))
End Function

Function WbAddDt(A As Workbook, Dt As Dt) As Worksheet
Dim O As Worksheet
Set O = WbAddWs(A, Dt.DtNm)
DrsPut DtDrs(Dt), WsA1(O)
Set WbAddDt = O
End Function

Private Sub DsWb__Tst()
Dim Wb As Workbook
Set Wb = DsWb(DsNew("Permit PermitD"))
WbVis Wb
Stop
Wb.Close False
End Sub

Sub Tst()
DsWb__Tst
End Sub
