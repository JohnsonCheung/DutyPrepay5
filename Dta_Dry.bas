Attribute VB_Name = "Dta_Dry"
Option Explicit
Option Compare Database

Sub DryAssertEq(A(), B())
If Not DryIsEq(A, B) Then Stop
End Sub

Function DryCol(Dry, Optional ColIdx% = 0) As Variant()
If AyIsEmpty(Dry) Then Exit Function
Dim O(), Dr
For Each Dr In Dry
    Push O, Dr(ColIdx)
Next
DryCol = O
End Function

Sub DryDmp(Dry)
AyDmp DryLy(Dry)
End Sub

Function DryIsEq(A(), B()) As Boolean
Dim N&: N = Sz(A)
If N <> Sz(B) Then Exit Function
If N = 0 Then DryIsEq = True: Exit Function
Dim J&, Dr
For Each Dr In A
    If Not AyIsEq(Dr, B(J)) Then Exit Function
    J = J + 1
Next
DryIsEq = True
End Function

Function DryNCol%(Dry)
Dim Dr, O%, M%
For Each Dr In Dry
    M = Sz(Dr)
    If M > O Then O = M
Next
DryNCol = O
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

Function DryStrCol(Dry, Optional ColIdx% = 0) As String()
DryStrCol = AySy(DryCol(Dry, ColIdx))
End Function
