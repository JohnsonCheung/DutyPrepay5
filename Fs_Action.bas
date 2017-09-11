Attribute VB_Name = "Fs_Action"
Option Explicit
Option Compare Database

Sub FilCpyToPth(FmFil, ToPth, Optional OvrWrt As Boolean)
Fso.CopyFile FmFil, ToPth & FfnFn(FmFil), OvrWrt
End Sub

Sub FtBrw(Ft)
Shell "code.cmd """ & Ft & """", vbHide
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

Function FtOpnInp%(Ft)
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

Sub PthRmvIfEmpty(P)
If Not IsPth(P) Then Exit Sub
If PthIsEmpty(P) Then Exit Sub
RmDir P
End Sub

Sub TmpPthBrw()
PthBrw TmpPth
End Sub
