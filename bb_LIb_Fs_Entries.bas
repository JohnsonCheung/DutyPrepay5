Attribute VB_Name = "bb_LIb_Fs_Entries"
Option Compare Database
Option Explicit
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
Function PthFfnAy(P, Optional Spec$ = "*.*") As String()
PthFfnAy = AddAyPfx(PthFnAy(P, Spec), P)
End Function
