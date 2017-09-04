Attribute VB_Name = "Xls_Put"
Option Explicit
Option Compare Database

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

Sub DrsPut(a As Drs, At As Range, Optional LoNm$)
AyPut a.Fny, At
SqPut RgRC(At, 2, 1), DrySq(a.Dry, Sz(a.Fny))
LoCrt RgWs(At), LoNm
End Sub

Function DrsWs(a As Drs, Optional WsNm$ = "Sheet1") As Worksheet
Dim O As Worksheet: Set O = WsNew(WsNm, Vis:=True)
DrsPut a, WsA1(O)
Set DrsWs = O
End Function

Sub DryPut(AtCell As Range, Dry)
AtCell.Value = DrySq(Dry)
End Sub

Function DsNDt%(a As Ds)
DsNDt = DtAySz(a.DtAy)
End Function

Function DsWb(a As Ds) As Workbook
Dim O As Workbook
Set O = WbNew
With WbFstWs(O)
    .Name = "Ds"
    .Range("A1").Value = a.DsNm
End With
If Not DsIsEmpty(a) Then
    Dim J%
    For J = 0 To DsNDt(a) - 1
        WbAddDt O, a.DtAy(J)
    Next
End If
Set DsWb = O
End Function

Function DtWs(a As Dt) As Worksheet
Dim O As Worksheet
Set O = WsNew(a.DtNm)
DrsPut DtDrs(a), WsA1(O)
Set DtWs = O
End Function

Sub SqPut(Cell As Range, Sq)
If SqIsEmpty(Sq) Then Exit Sub
ReSzRg(Cell, Sq).Value = Sq
End Sub

Function TblWs(T, Optional D As Database) As Worksheet
Set TblWs = DtWs(TblDt(T, D))
End Function

Function WbAddDt(a As Workbook, Dt As Dt) As Worksheet
Dim O As Worksheet
Set O = WbAddWs(a, Dt.DtNm)
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

