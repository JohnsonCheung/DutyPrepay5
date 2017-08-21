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
Private Sq_()
Property Get Sq() As Variant()
Sq = Sq_
End Property
Function Init(Sq()) As Sq
Dim I&:
I = UBound(Sq, 2)
I = UBound(Sq, 1)
Sq_ = Sq
Set Init = Me
End Function

Sub PutAt(At As Range)
Rg(At).ReSz(Sq_).Value = Sq_
End Sub

Function Dr(R&, Optional CnoAy) As Variant()
Dim mCnoAy%()
    Dim J%
    If IsMissing(CnoAy) Then
        ReDim mCnoAy(UBound(Sq_, 2) - 1)
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
        O(J) = Sq_(R, C)
    Next
Dr = O
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
Function InitByVAy(VAy) As Sq
If AyIsEmpty(VAy) Then Exit Function
Dim O(), R&
ReDim O(1 To 1, 1 To Sz(VAy))
R = 0
Dim V
For Each V In VAy
    R = R + 1
    O(R, 1) = V
Next
Set InitByVAy = Init(O)
End Function


Function Sel(Optional MapStr$) As Drs
Dim Fny$(), Fm$() 'MapStr
    If MapStr = "" Then
        Fny = AySy(Dr(1))
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
    A = Ay(Dr(1)).IdxAy(Fm)
    U = UB(A)
    ReDim CnoAy(U)
    For J = 0 To U
        CnoAy(J) = A(J) + 1
    Next
Dim Dry As Dry 'Sq,SqIdxAy
    Dim R&, Cno%, C%
    Dim UFld%
    Dim Idx%
    UFld = UB(CnoAy)
    For R = 2 To NRow
        ReDim Dr(UFld)
        For C = 0 To UFld
            Cno = CnoAy(C)
            If Cno > 0 Then
                Dr(C) = Sq_(R, Cno)
            End If
        Next
        Dry.AddDr Dr
    Next
Set Sel = Nw.Drs(Fny, Dry)
End Function

Function NRow&()
NRow = UBound(Sq_, 1)
End Function

Function NCol&()
NCol = UBound(Sq_, 2)
End Function

