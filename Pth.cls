VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Private P$
Property Get Pth$()
Pth = P
End Property
Function Init(Pth) As Pth
P = Pth
Set Init = Me
End Function
Property Let Pth(Pth_$)
If LasChr(Pth_) <> "\" Then Err.Raise 1, , FmtQQ("LasChr of given path must be [\].  Given Path = [?]", Pth_)
P = Pth
End Property
Function FfnAy(Optional Spec$ = "*.*") As String()
FfnAy = Ay(FnAy(Spec)).AddPfx(P)
End Function

Function FnAy(Optional Spec$ = "*.*") As String()
Dim O$()
Dim M$
M = Dir(P & Spec)
While M <> ""
    Push O, M
    M = Dir
Wend
FnAy = O
End Function
Sub Ens()
If IsExist Then Exit Sub
MkDir P
End Sub


Function Brw() As Pth
Shell "Explorer """ & P & """", vbMaximizedFocus
Set Brw = Me
End Function

Sub AsstIsExist()
If Not IsExist Then Err.Raise 1, , FmtQQ("Given Pth[?] does not exist", P)
End Sub

Sub AsrtSfx()
If LasChr(P) <> "\" Then Err.Raise 1, , FmtQQ("Given Pth[?] does not end with \", P)
End Sub

Function IsExist() As Boolean
IsExist = Dir(P, vbDirectory) <> ""
End Function

Function HasFil() As Boolean
AsrtSfx
If Not IsExist Then Exit Function
HasFil = (Dir(P & "*.*") <> "")
End Function

Function HasSubDir() As Boolean
If Not IsExist Then Exit Function
AsrtSfx
Dim A$: A = Dir(P & "*.*", vbDirectory)
Dir
HasSubDir = Dir <> ""
End Function

Function IsEmpty()
AsrtSfx
If HasFil Then Exit Function
If HasSubDir Then Exit Function
IsEmpty = True
End Function

Function PthAy(Optional Spec$ = "*.*") As String()
PthAy = Ay(SubDirAy(Spec)).AddPfx(P)
End Function
Function ClrFil() As Pth
If Not IsExist Then Exit Function
Dim Ay$(): Ay = FfnAy
Dim F
On Error Resume Next
For Each F In Ay
    Kill F
Next
End Function

Sub RmvEmptySubDir()
If Not IsExist Then Exit Sub
Dim A$(): A = PthAy: If AyIsEmpty(A) Then Exit Sub
Dim I
For Each I In A
    Nw.Pth(I).RmvIfEmpty
Next
End Sub

Sub RmvIfEmpty()
If Not IsExist Then Exit Sub
If IsEmpty Then RmDir P
End Sub

Function SubDirAy(Optional Spec$ = "*.*") As String()
AsrtSfx
Dir P & Spec, vbDirectory
Dir
Dim A$, O$()
A = Dir
While A <> ""
    Push O, A
    A = Dir
Wend
SubDirAy = O
End Function
