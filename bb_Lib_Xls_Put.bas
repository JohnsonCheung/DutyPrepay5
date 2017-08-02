Attribute VB_Name = "bb_Lib_Xls_Put"
Option Compare Database
Option Explicit
Function DtWs(Dt As Dt) As Worksheet
Dim O As Worksheet
Set O = NewWs
PutDt Dt, O
Set DtWs = O
End Function
Sub PutDt(A As Dt, Ws As Worksheet)
PutAy WsA1(Ws), A.Fny
PutSq WsRC(Ws, 2, 1), DrySq(A.Dry, Sz(A.Fny))
End Sub
Function AddDtToWb(A As Dt, Wb As Workbook) As Worksheet
Dim O As Worksheet
Set O = AddWs(Wb, A.DtNm)
PutDt A, O
Set AddDtToWb = O
End Function
Sub AddDsToWb__Tst()
Dim Ds As Ds, Wb As Workbook
Ds = NewDs("Permit PermitD")
Set Wb = NewWb
AddDsToWb Ds, Wb
WbVis Wb
Stop
Wb.Close False
End Sub
Sub PutSq(Cell As Range, Sq)
ReSzRg(Cell, Sq).Value = Sq
End Sub
Sub PutAy(Cell As Range, Ay)
PutSq Cell, AySqH(Ay)
End Sub
Function TblWs(T, Optional D As Database) As Worksheet
Set TblWs = DtWs(TblDt(T, D))
End Function

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
Sub PutDry(AtCell As Range, Dry)
AtCell.Value = DrySq(Dry)
End Sub
Sub AddDsToWb(A As Ds, Wb As Workbook)
Dim J%
For J = 0 To DtAySz(A.DtAy) - 1
    AddDtToWb A.DtAy(J), Wb
Next
End Sub
Function DryNCol%(Dry)
Dim Dr, O%, M%
For Each Dr In Dry
    M = Sz(Dr)
    If M > O Then O = M
Next
DryNCol = M
End Function
Function DrySq(Dry, Optional NCol% = 0) As Variant()
If IsEmptyAy(Dry) Then Exit Function
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

