Attribute VB_Name = "Fs_Entries"
Option Explicit
Option Compare Database

Sub AssertIsPth(P)
Const CSub$ = "AssertIsPth"
If Not IsPth(P) Then Er CSub, "Given {Pth} is not ends with \ or not exists", P
End Sub

Function IsPth(P) As Boolean
If LasChr(P) <> "\" Then Exit Function
IsPth = Fso.FolderExists(P)
End Function

Function PthFfnAy(P, Optional Spec$ = "*.*") As String()
PthFfnAy = AyAddPfx(PthFnAy(P, Spec), P)
End Function

Function PthFnAy(P, Optional Spec$ = "*.*") As String()
AssertIsPth P
Dim O$()
Dim M$
M = Dir(P & Spec)
While M <> ""
    Push O, M
    M = Dir
Wend
PthFnAy = O
End Function
