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

Function DryMge(Dry, MgeIdx%, Sep$) As Variant()
Dim O(), J%
Dim Idx%
For J = 0 To UB(Dry)
    Idx = DryMgeIdx(O, Dry(J), MgeIdx)
    If Idx = -1 Then
        Push O, Dry(J)
    Else
        O(Idx)(MgeIdx) = O(Idx)(MgeIdx) & Sep & Dry(J)(MgeIdx)
    End If
Next
DryMge = O
End Function

Function DryMgeIdx&(Dry, Dr, MgeIdx%)
Dim O&, D, J%
For O = 0 To UB(Dry)
    D = Dry(O)
    For J = 0 To UB(Dr)
        If J <> MgeIdx Then
            If Dr(J) <> D(J) Then GoTo Nxt
        End If
    Next
    DryMgeIdx = O
    Exit Function
Nxt:
Next
DryMgeIdx = -1
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

Function TblLyFmt(TblLy) As String()
AyAssertPfx TblLy, "|"
Dim Dry()
    Dim I
    For Each I In TblLy
        Push Dry, AyTrim(SplitVBar(I))
    Next
TblLyFmt = DryLy(Dry)
End Function

Sub TblLyFmt__Tst()
Dim TblLy$()
Dim Act$()
Dim Exp$()
Push TblLy, "|lskdf|sdlf|lsdkf"
Push TblLy, "|lsdf|"
Push TblLy, "|lskdfj|sdlfk|sdlkfj|sdklf|skldf|"
Push TblLy, "|sdf"
Act = TblLyFmt(TblLy)
Exp = Sy()
AyDmp Act
AyAssertEq Exp, Act
End Sub
