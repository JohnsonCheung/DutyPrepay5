VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BrkLin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Private A_$() 'Either DryLy or DrsLy
Private Hdr_$
Private ColIdx_%
Private ColNm_$
Private IsDrs_ As Boolean
Function DryLy_() As String()
DryLy_ = A_
End Function
Function InitByDryLy(DryLy$(), ColIdx%) As BrkLin
If Sz(DryLy) < 2 Then Stop
ColIdx_ = ColIdx
A_ = DryLy
Nw.Ay(A_).RmvAtCnt 0, 2
Set InitByDryLy = Me
End Function
Function InitByDrsLy(DrsLy$(), ColNm$) As BrkLin
If Sz(DrsLy) < 3 Then Stop
Hdr_ = DrsLy(1)
ColNm_ = ColNm
IsDrs_ = True
Dim Fny$(): Fny = SplitVBar(RmvLasChr(RmvFstChr(Hdr_)))
ColIdx_ = Ay(Fny).Trim.Idx(ColNm_)
A_ = DrsLy
Nw.Ay(A_).RmvAtCnt 0, 2
Set InitByDrsLy = Me
End Function
Sub Brw()
LyX.Brw
End Sub
Property Get LyX() As Ay
Dim O As New Ay
Set LyX = O.Init(Ly)
End Property
Function Ly() As String()
If ColIdx_ = -1 Then Ly = A_: Exit Function
If Sz(A_) = 2 Then Exit Function
Dim Fm&, L%
    Dim N%: N = ColIdx_ + 1
    Dim P1&, P2&
    P1 = InstrN(Hdr_, "|", N)
    P2 = InStr(P1 + 1, Hdr_, "|")
    Fm = P1 + 1
    L = P2 - P1 - 1
Dim O$()
    Push O, A_(0)
    Dim LasV$: LasV = Mid(A_(1), Fm, L)
    Dim J&
    Dim V$
    For J = 1 To UB(A_) - 1
        V = Mid(A_(J), Fm, L)
        If LasV <> V Then
            Push O, Hdr_
            LasV = V
        End If
        Push O, A_(J)
    Next
    Push O, A_(0)
Ly = O
End Function
