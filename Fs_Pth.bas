Attribute VB_Name = "Fs_Pth"
Option Explicit
Option Compare Database

Sub PthAssertIsExist(P)
If Not IsPth(P) Then Err.Raise 1, , FmtQQ("Given Pth[?] does not exist", P)
End Sub

Sub PthAssertSfx(P)
If LasChr(P) <> "\" Then Err.Raise 1, , FmtQQ("Given Pth[?] does not end with \", P)
End Sub

Function PthHasFil(P) As Boolean
PthAssertSfx P
If Not IsPth(P) Then Exit Function
PthHasFil = (Dir(P & "*.*") <> "")
End Function

Function PthHasSubDir(P) As Boolean
If Not IsPth(P) Then Exit Function
PthAssertSfx P
Dim A$: A = Dir(P & "*.*", vbDirectory)
Dir
PthHasSubDir = Dir <> ""
End Function

Function PthIsEmpty(P)
PthAssertIsExist P
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
PthAssertSfx P
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
