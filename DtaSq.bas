Attribute VB_Name = "DtaSq"
Option Compare Database
Option Explicit

Function SqDr(Sq, R&, Optional CnoAy) As Variant()
Dim mCnoAy%()
    Dim J%
    If IsMissing(CnoAy) Then
        ReDim mCnoAy(UBound(Sq, 2) - 1)
        For J = 0 To UB(mCnoAy)
            mCnoAy(J) = J + 1
        Next
    Else
        mCnoAy = CnoAy
    End If
Dim UCol%
    UCol = UB(mCnoAy)
Dim O()
    ReDim O(UCol)
    Dim C%
    For J = 0 To UCol
        C = mCnoAy(J)
        O(J) = Sq(R, C)
    Next
SqDr = O
End Function

Function SqIsEmpty(Sq) As Boolean
SqIsEmpty = True
On Error GoTo X
Dim A
If UBound(Sq, 1) < 0 Then Exit Function
If UBound(Sq, 2) < 0 Then Exit Function
SqIsEmpty = False
Exit Function
X:
End Function

Function SqSel(Sq, Optional MapStr$) As Drs
Dim Fny$(), Fm$() 'MapStr
    If MapStr = "" Then
        Fny = AyStrAy(SqDr(Sq, 1))
        Fm = Fny
    Else
        With BrkMapStr(MapStr)
            Fny = .Sy1
            Fm = .Sy2
        End With
    End If
Dim SqCnoAy%() 'Fm,Sq
    Dim A&()
    Dim U%
    Dim J%
    A = AyIdxAy(SqDr(Sq, 1), Fm)
    U = UB(A)
    ReDim SqCnoAy(U)
    For J = 0 To U
        SqCnoAy(J) = A(J) + 1
    Next
Dim Dry() 'Sq,SqIdxAy
    Dim R&, Cno%, C%
    Dim UFld%
    Dim Idx%
    Dim Dr()
    UFld = UB(SqCnoAy)
    For R = 2 To UBound(Sq, 1)
        ReDim Dr(UFld)
        For C = 0 To UFld
            Cno = SqCnoAy(C)
            If Cno > 0 Then
                Dr(C) = Sq(R, Cno)
            End If
        Next
        Push Dry, Dr
    Next
SqSel.Dry = Dry
SqSel.Fny = Fny
End Function
