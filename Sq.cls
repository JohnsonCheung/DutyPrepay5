VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Private Sq()
Function Init(Sq_()) As Sq
Dim I&:
I = UBound(Sq_, 2)
I = UBound(Sq_, 1)
Sq = Sq_
Set Init = Me
End Function

Function Dr(R&, Optional CnoAy) As Variant()
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
Dr = O
End Function
Function Xls() As SqXls
Set Xls = SqXls(Me)
End Function
Function InitByHAy(HAy) As Sq
If AyIsEmpty(HAy) Then Exit Function
Dim O(), C%
ReDim O(1 To 1, 1 To Sz(HAy))
C = 0
Dim V
For Each V In HAy
    C = C + 1
    O(1, C) = V
Next
Set InitByHAy = Init(O)
End Function


Function Sel(Optional MapStr$) As Drs
Dim Fny$(), Fm$() 'MapStr
    If MapStr = "" Then
        Fny = AyStrAy(Dr(1))
        Fm = Fny
    Else
        With BrkMapStr(MapStr)
            Fny = .Sy1
            Fm = .Sy2
        End With
    End If
Dim CnoAy%() 'Fm,Sq
    Dim A&()
    Dim U%
    Dim J%
    A = AyIdxAy(Dr(1), Fm)
    U = UB(A)
    ReDim SqCnoAy(U)
    For J = 0 To U
        SqCnoAy(J) = A(J) + 1
    Next
Dim Dry As Dry 'Sq,SqIdxAy
    Dim R&, Cno%, C%
    Dim UFld%
    Dim Idx%
    UFld = UB(CnoAy)
    For R = 2 To UBound(Sq, 1)
        ReDim Dr(UFld)
        For C = 0 To UFld
            Cno = CnoAy(C)
            If Cno > 0 Then
                Dr(C) = Sq(R, Cno)
            End If
        Next
        Dry.Push Dr
    Next
Set Sel = ccNew.Drs(Fny, Dry)
End Function

