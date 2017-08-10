Attribute VB_Name = "bb_Lib_Fs_Action"
Option Compare Database
Option Explicit

Sub FilCpyToPth(FmFil, ToPth, Optional OvrWrt As Boolean)
Fso.CopyFile FmFil, ToPth & FfnFn(FmFil), OvrWrt
End Sub

Sub FtBrw(Ft)
Shell "NotePad """ & Ft & """", vbMaximizedFocus
End Sub

Function FtLy(Ft) As String()
Dim F%: F = FtOpnInp(Ft)
Dim L$, O$()
While Not EOF(F)
    Line Input #F, L
    Push O, L
Wend
Close #F
FtLy = O
End Function

Function FtOpnApp%(Ft)
Dim O%: O = FreeFile(1)
Open Ft For Append As #O
FtOpnApp = O
End Function

Function FtOpnInp(Ft)
Dim O%: O = FreeFile(1)
Open Ft For Input As #O
FtOpnInp = O
End Function

Function FtOpnOup%(Ft)
Dim O%: O = FreeFile(1)
Open Ft For Output As #O
FtOpnOup = O
End Function

Sub PthBrw(P)
Shell "Explorer """ & P & """", vbMaximizedFocus
End Sub

Sub PthClrFil(P)
If Not IsPth(P) Then Exit Sub
Dim Ay$(): Ay = PthFfnAy(P)
Dim F
On Error Resume Next
For Each F In Ay
    Kill F
Next
End Sub

Sub PthEns(P)
If IsPth(P) Then Exit Sub
MkDir P
End Sub

Sub RmvPthIfEmpty(P)
If Not IsPth(P) Then Exit Sub
If IsEmptyPth(P) Then Exit Sub
RmDir P
End Sub

Sub TmpPthBrw()
PthBrw TmpPth
End Sub
