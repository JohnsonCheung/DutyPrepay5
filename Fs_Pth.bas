Attribute VB_Name = "Fs_Pth"
Option Compare Database
Option Explicit

Sub AssertIsPth(P)
If Not IsPth(P) Then Err.Raise 1, , FmtQQ("Given Pth[?] does not exist", P)
End Sub

Sub AssertPth(P)
If LasChr(P) <> "\" Then Err.Raise 1, , FmtQQ("Given Pth[?] does not end with \", P)
End Sub

Function IsEmptyPth(P) As Boolean
AssertIsPth P
If PthHasSubDir(P) Then Exit Function
End Function

Function IsPth(P) As Boolean
IsPth = Dir(P, vbDirectory) <> ""
End Function

Function PthHasFil(P) As Boolean
AssertPth P
If Not IsPth(P) Then Exit Function
PthHasFil = (Dir(P & "*.*") <> "")
End Function

Function PthHasSubDir(P) As Boolean
If Not IsPth(P) Then Exit Function
AssertPth P
Dim A$: A = Dir(P & "*.*", vbDirectory)
Dir
PthHasSubDir = Dir <> ""
End Function

Function PthIsEmpty(P)
If PthHasFil(P) Then Exit Function
If PthHasSubDir(P) Then Exit Function
PthIsEmpty = True
End Function

Function PthPthAy(P, Optional Spec$ = "*.*") As String()
PthPthAy = AyAddPfx(PthSubDirAy(P, Spec), P)
End Function

Sub PthRmvEmptySubDir(P)
Dim A$(): A = PthPthAy(P): If AyIsEmpty(A) Then Exit Sub
Dim I
For Each I In A
    PthRmvIfEmpty I
Next
End Sub

Sub PthRmvIfEmpty(P)
If PthIsEmpty(P) Then RmDir P
End Sub

Function PthSubDirAy(P, Optional Spec$ = "*.*") As String()
AssertPth P
Dir P & Spec, vbDirectory
Dir
Dim A$, O$()
A = Dir
While A <> ""
    Push O, A
    A = Dir
Wend
PthSubDirAy = O
End Function

Private Sub PthRmvEmptySubDir__Tst()
TmpPth
End Sub

Sub Tst()
PthRmvEmptySubDir__Tst
End Sub
