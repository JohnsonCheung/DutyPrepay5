VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ffn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Public Ffn$
Function Init(Ffn_) As Ffn
Ffn = Ffn_
Set Init = Me
End Function

Function AddFnSfx$(Sfx)
AddFnSfx = RmvExt & Sfx & Ext
End Function
Function AddFnSfxX(Sfx) As Ffn
Set AddFnSfxX = Nw.Ffn(AddFnSfx(Sfx))
End Function

Function Ext$()
Dim P%: P = InStrRev(Ffn, ".")
If P = 0 Then Exit Function
Ext = Mid(Ffn, P)
End Function

Function Fn$()
Dim P%: P = InStrRev(Ffn, "\")
If P = 0 Then Fn = Ffn: Exit Function
Fn = Mid(Ffn, P + 1)
End Function
Sub CpyToPth(ToPth, Optional OvrWrt As Boolean)
AsrtIsPth ToPth
Fso.CopyFile Ffn, ToPth & Fn, OvrWrt
End Sub

Function Fnn$()
Fnn = RmvExt
End Function

Function Pth$()
Dim P%: P = InStrRev(Ffn, "\")
If P = 0 Then Exit Function
Pth = Left(Ffn, P)
End Function

Function RmvExt$()
Dim P%: P = InStrRev(Ffn, ".")
If P = 0 Then RmvExt = Left(Ffn, P): Exit Function
RmvExt = Left(Ffn, P - 1)
End Function

Function RplExt$(NewExt)
RplExt = RmvExt & NewExt
End Function


