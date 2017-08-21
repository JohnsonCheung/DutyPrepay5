VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tst_MdSrc_SrtLy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Sub Run()
Dim Pj As Pj
Set Pj = Nw.Pj(CurPj)
Dim I
For Each I In Pj.MdAy
    TstMd I
Next
End Sub
Private Sub TstMd(Md)
Dim M As Md: Set M = Md
Dim Srt$(), Ly$()
Ly = M.Ly
Srt = M.Src.SrtLy
AssertIsAySam Srt, Ly
End Sub
Private Sub AssertIsAySam(Ay1, Ay2)
AssertNotEmpty AyMinus(Ay1, Ay2)
AssertNotEmpty AyMinus(Ay2, Ay1)
End Sub
Private Sub AssertNotEmpty(Ay)
If Not AyIsEmpty(Ay) Then Stop
End Sub
