Attribute VB_Name = "Ide_Enm"
Option Explicit
Option Compare Database
Private Enum aa 'For testing
    AA1
    '
    
End Enum

Function EnmBdyLy(EnmNm$, Optional A As CodeModule) As String()
Dim B%: B = EnmLinIdx(EnmNm, A): If B = -1 Then Exit Function
Dim O$(), Ly$(), J%
Ly = MdDclLy(A)
For J = B To UB(Ly)
    Push O, Ly(J)
    If IsPfx(Ly(J), "End Enum") Then EnmBdyLy = O: Exit Function
Next
Stop
End Function

Function EnmIsMbrLin(L) As Boolean
If SrcLinIsRmk(L) Then Exit Function
If Trim(L) = "" Then Exit Function
EnmIsMbrLin = True
End Function

Function EnmLinIdx%(EnmNm$, Optional A As CodeModule)
Dim Ly$(): Ly = MdDclLy(A)
Dim U%: U = UB(Ly)
Dim O%, L$
For O = 0 To U
    If SrcLinIsEnm(Ly(O)) Then
        L = Ly(O)
        ParseMdy L
        L = RmvFstTerm(L)
        If FstTerm(L) = EnmNm Then
            EnmLinIdx = O: Exit Function
        End If
    End If
Next
EnmLinIdx = -1
End Function

Function EnmMbrCnt%(EnmNm$, Optional A As CodeModule)
EnmMbrCnt = Sz(EnmMbrLy(EnmNm, A))
End Function

Function EnmMbrLy(EnmNm$, Optional A As CodeModule) As String()
Dim Ly$(), O$(), J%
Ly = EnmBdyLy(EnmNm, A)
For J = 1 To UB(Ly) - 1
    If EnmIsMbrLin(Ly(J)) Then Push O, Ly(J)
Next
EnmMbrLy = O
End Function

Private Sub EnmBdyLy__Tst()
AyDmp EnmBdyLy("AA")
End Sub

Private Sub EnmLno__Tst()
Debug.Assert EnmLinIdx("AA", Md("Ide")) = 2
End Sub

Private Sub EnmMbrCnt__Tst()
Debug.Assert EnmMbrCnt("AA", Md("Ide")) = 1
End Sub
