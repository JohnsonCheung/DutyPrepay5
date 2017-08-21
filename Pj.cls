VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Private Pj As VBProject
Function Init(Optional A As VBProject) As Pj
Set Pj = DftPj(A)
Set Init = Me
End Function
Sub Tst()

End Sub
Sub Exp()
CpyToSrc
SrcPthX.ClrFil
Dim Md As Md, I
For Each I In MdAy
    Set Md = I
    Md.Exp
Next
End Sub
Function SrcPthX() As Pth
Set SrcPthX = Pth(SrcPth)
End Function
Function CrtMd(MdNm$, Ty As vbext_ComponentType) As Md
Dim O As VBComponent
Set O = Pj.VBComponents.Add(Ty)
O.Name = MdNm
Set CrtMd = Nw.Md(O.CodeModule)
End Function
Function FunDrs(Optional WithBdyLy As Boolean, Optional WithBdyLines As Boolean, Optional WithRmkLines As Boolean) As Drs
Dim Dry As New Dry
    Dim I, Md As Md
    For Each I In MdAy
        Set Md = I
        Dry.AddDrAy Md.Src.FunDrs(WithBdyLy, WithBdyLines).Dry.Dry
    Next
Dim Fny$()
'    Fny = Md.Src.FunDrsFny(WithBdyLy, WithBdyLines, WithRmkLines)
Set FunDrs = Nw.Drs(Fny, Dry)
End Function

Sub CpyToSrc()
Ffn(Pj.FileName).CpyToPth SrcPth, OvrWrt:=True
End Sub

Function MdAy() As Md()
Dim O() As Md
Dim Cmp As VBComponent
For Each Cmp In Pj.VBComponents
    Push O, Nw.Md(Cmp.CodeModule)
Next
MdAy = O
End Function
Function Md(MdNm$) As Md
Set Md = Nw.Md(Pj.VBComponents(MdNm).CodeModule)
End Function
Private Sub FunDrs__Tst()
Dim Drs As Drs: Set Drs = ZZ.FunDrs(WithBdyLines:=True)
Drs.WsX Nm, Vis:=True
End Sub
Private Function ZZ() As Pj
Set ZZ = Nw.Pj
End Function
Function Nm$()
Nm = Pj.Name
End Function
Private Sub MdAy__Tst()
Dim O() As Md
O = Pj.MdAy
Dim I, Md As Md
For Each I In O
    Set Md = I
    Debug.Print Md.Nm
Next
End Sub
Function PjFfn() As Ffn
Set PjFfn = Ffn(PjFilNm)
End Function
Function PjFilNm$()
PjFilNm = Pj.FileName
End Function
Function SrcPth$()
Dim O$:
With PjFfn
O = .Pth & "Src\": Pth(O).Ens
O = O & .Fn & "\": Pth(O).Ens
End With
SrcPth = O
End Function

Sub Srt()
Dim Md As Md, I
For Each I In MdAy
    Set Md = I
    Md.Srt
Next
End Sub

Function TstFunNy_WithEr() As String()
Dim O$(), I, Md As Md
Dim P$
P = Md.Nm & "."
For Each I In MdAy
    Set Md = I
    PushAy O, Ay(Md.Src.TstFunNy_WithEr).AddPfx(P)
    Next
TstFunNy_WithEr = O
End Function

Sub UpdTstFun()
Dim I, Md As Md
For Each I In MdAy
    Set Md = I
    Md.UpdTstFun
Next
End Sub
