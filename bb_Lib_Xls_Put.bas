Attribute VB_Name = "bb_Lib_Xls_Put"
Option Compare Database
Option Explicit
Function DtWs(A As Dt) As Worksheet
Dim O As Worksheet
Set O = WsNew(A.DtNm)
DrsPut DtDrs(A), WsA1(O)
Set DtWs = O
End Function
Function DtDrs(A As Dt) As Drs
Dim O As Drs
O.Fny = A.Fny
O.Dry = A.Dry
DtDrs = O
End Function
Sub DrsPut(A As Drs, At As Range, Optional LoNm$)
AyPut A.Fny, At
SqPut RgRC(At, 2, 1), DrySq(A.Dry, Sz(A.Fny))
LoNew RgWs(At), LoNm
End Sub
Function WbAddDt(A As Dt, Wb As Workbook) As Worksheet
Dim O As Worksheet
Set O = AddWs(Wb, A.DtNm)
DrsPut DtDrs(A), WsA1(O)
Set WbAddDt = O
End Function
Private Sub WbAddDs__Tst()
Dim Ds As Ds, Wb As Workbook
Ds = DsNew("Permit PermitD")
Set Wb = NewWb
WbAddDs Ds, Wb
WbVis Wb
Stop
Wb.Close False
End Sub
Sub SqPut(Cell As Range, Sq)
ReSzRg(Cell, Sq).Value = Sq
End Sub
Function DrsWs(A As Drs, Optional WsNm$ = "Sheet1") As Worksheet
Dim O As Worksheet: Set O = WsNew(WsNm, Vis:=True)
DrsPut A, WsA1(O)
Set DrsWs = O
End Function
Sub AyPut(Ay, Cell As Range)
SqPut Cell, AySqH(Ay)
End Sub
Function TblWs(T, Optional D As Database) As Worksheet
Set TblWs = DtWs(TblDt(T, D))
End Function

Sub DryPut(AtCell As Range, Dry)
AtCell.Value = DrySq(Dry)
End Sub
Sub WbAddDs(A As Ds, Wb As Workbook)
Dim J%
For J = 0 To DtAySz(A.DtAy) - 1
    WbAddDt A.DtAy(J), Wb
Next
End Sub

Function AySqH(Ay) As Variant()
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
Function DryNCol%(Dry)
Dim Dr, O%, M%
For Each Dr In Dry
    M = Sz(Dr)
    If M > O Then O = M
Next
DryNCol = M
End Function
Function DrySq(Dry, Optional NCol% = 0) As Variant()
If AyIsEmpty(Dry) Then Exit Function
Dim NRow&
    If NCol = 0 Then NCol = DryNCol(Dry)
    NRow = Sz(Dry)
Dim O()
    ReDim O(1 To NRow, 1 To NCol)
Dim C%, R&, Dr
    R = 0
    For Each Dr In Dry
        R = R + 1
        For C = 0 To UB(Dr)
            O(R, C + 1) = Dr(C)
        Next
    Next
DrySq = O
End Function

Sub Tst()
WbAddDs__Tst
End Sub
