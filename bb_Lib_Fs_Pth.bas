Attribute VB_Name = "bb_Lib_Fs_Pth"
Option Compare Database
Option Explicit
Sub EnsPth(P)
If IsPth(P) Then Exit Sub
MkDir P
End Sub
Function IsPth(P) As Boolean
IsPth = Dir(P, vbDirectory) <> ""
End Function
Sub BrwPth(P)
Shell "Explorer """ & P & """", vbMaximizedFocus
End Sub
Sub ClrPthFil(P)
If Not IsPth(P) Then Exit Sub
Dim Ay$(): Ay = PthFfnAy(P)
Dim F
On Error Resume Next
For Each F In Ay
    Kill F
Next
End Sub
