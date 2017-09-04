Attribute VB_Name = "Fs_Entries"
Option Explicit
Option Compare Database

Function PthFfnAy(P, Optional Spec$ = "*.*") As String()
PthFfnAy = AyAddPfx(PthFnAy(P, Spec), P)
End Function

Function PthFnAy(P, Optional Spec$ = "*.*") As String()
Dim O$()
Dim M$
M = Dir(P & Spec)
While M <> ""
    Push O, M
    M = Dir
Wend
PthFnAy = O
End Function
