Attribute VB_Name = "bb_Lib_Fs_Pth"
Option Compare Database
Option Explicit
Function IsPth(P) As Boolean
IsPth = Dir(P, vbDirectory) <> ""
End Function
Function HasSubDir(P) As Boolean
If Not IsPth(P) Then Exit Function
AssertPth P
Dim A$: A = Dir(P & "*.*", vbDirectory)
Dir
HasSubDir = Dir <> ""
End Function
Sub AssertPth(P)
If LasChr(P) <> "\" Then Err.Raise 1, , FmtQQ("Given Pth[?] does not end with \", P)
End Sub
Function HasFil(P) As Boolean
AssertPth P
If Not IsPth(P) Then Exit Function
HasFil = (Dir(P & "*.*") <> "")
End Function
Sub AssertIsPth(P)
If Not IsPth(P) Then Err.Raise 1, , FmtQQ("Given Pth[?] does not exist", P)
End Sub
Function IsEmptyPth(P) As Boolean
AssertIsPth P
If HasSubDir(P) Then Exit Function
End Function
