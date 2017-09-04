Attribute VB_Name = "Fs_Pth"
Option Explicit
Option Compare Database

Function IsPth(P) As Boolean
IsPth = Dir(P, vbDirectory) <> ""
End Function

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
Dim a$: a = Dir(P & "*.*", vbDirectory)
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
Dim a$(): a = PthPthAy(P): If AyIsEmpty(a) Then Exit Sub
Dim I
For Each I In a
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
Dim a$, O$()
a = Dir
While a <> ""
    Push O, a
    a = Dir
Wend
PthSubDirAy = O
End Function

Private Sub PthRmvEmptySubDir__Tst()
TmpPth
End Sub

Sub Tst()
PthRmvEmptySubDir__Tst
End Sub
