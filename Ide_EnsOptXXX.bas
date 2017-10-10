Attribute VB_Name = "Ide_EnsOptXXX"
Option Explicit
Option Compare Database

Function MdAddOptXXX(OptXXX$, Optional A As CodeModule)
DftMd(A).InsertLines 1, "Option " & OptXXX
Debug.Print MdNm(A), "<-- Option " & OptXXX & " added"
End Function

Sub MdEnsOptCmpDb(Optional A As CodeModule)
MdEnsOptXXX "Compare Database", A
End Sub

Sub MdEnsOptExplicit(Optional A As CodeModule)
MdEnsOptXXX "Explicit", A
End Sub

Sub MdEnsOptXXX(OptXXX$, Optional A As CodeModule)
Const CSub$ = "MdEnsOptXXX"
If OptXXX = "Explicit" And OptXXX <> "Compare Database" Then Er CSub, "OptXXX must be [Explicit] or [Compare Database]", OptXXX
If MdHasOptXXX(OptXXX, A) Then
    Debug.Print MdNm(A), "(* With Option Explicit *)"
Else
    Debug.Print MdNm(A), "<-------------------- No Option Explicit"
    MdAddOptXXX OptXXX, A
End If

End Sub

Function MdHasOptXXX(OptXXX$, Optional A As CodeModule) As Boolean
MdHasOptXXX = SrcHasOptXXX(OptXXX, MdDclLy(A))
End Function

Function SrcEnsOptCmpDb(Src$()) As String()
SrcEnsOptCmpDb = SrcEnsOptXXX("Compare Database", Src)
End Function

Function SrcEnsOptExplicit(Src$()) As String()
SrcEnsOptExplicit = SrcEnsOptXXX("Explicit", Src)
End Function

Function SrcEnsOptXXX(OptXXX$, Src$()) As String()
If SrcHasOptXXX(OptXXX, Src) Then
    SrcEnsOptXXX = Src
    Debug.Print "Src (* With Option Explicit *)"
Else
    Debug.Print "Src <-------------------- No Option " & OptXXX
    SrcEnsOptXXX = AyIns(Src, "Option " & OptXXX)
End If
End Function

Function SrcHasOptXXX(OptXXX$, Src$()) As Boolean
Dim Ay$()
    Ay = SrcDclLy(Src)
If AyIsEmpty(Ay) Then Exit Function
Dim I
For Each I In Ay
    If I = "Option " & OptXXX Then SrcHasOptXXX = True: Exit Function
Next
End Function
