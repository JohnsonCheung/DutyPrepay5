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

Sub Exp()
CpyToSrc
PthClrFil SrcPth
Dim Md As Md, I
For Each I In MdAy
    Set Md = I
    Md.Exp
Next
End Sub
Function CrtMd(MdNm$, Ty As vbext_ComponentType) As Md
Dim O As VBComponent
Set O = Pj.VBComponents.Add(Ty)
O.Name = MdNm
Set CrtMd = ccNew.Md(O.CodeModule)
End Function
Function FunDrs(Optional WithBdyLy As Boolean, Optional WithBdyLines As Boolean) As Drs
Dim Dry As New Dry
    Dim I, Md As Md
    For Each I In MdAy
        Set Md = I
        Dry.PushAy Md.FunDrs(WithBdyLy, WithBdyLines).Dry
    Next
Dim Fny$()
    Fny = Md.FunDrsFny(WithBdyLy, WithBdyLines)
FunDrs = ccNew.Drs(Fny, Dry)
End Function
Sub CpyToSrc()
FilCpyToPth Pj.FileName, SrcPth, OvrWrt:=True
End Sub
Function MdAy() As Md()
Dim O() As Md
Dim Cmp As VBComponent
For Each Cmp In Pj.VBComponents
    PushObj O, ccNew.Md(Cmp.CodeModule)
Next
MdAy = O
End Function
Function Md(MdNm$) As Md
Set Md = ccNew.Md(Pj.VBComponents(MdNm).CodeModule)
End Function
Private Sub FunDrs__Tst()
Dim Drs As Drs: Set Drs = FunDrs(WithBdyLines:=True)
DrsXls(Drs).Ws Nm, Vis:=True
End Sub

Function Nm$()
Nm = Pj.Name
End Function
Private Sub MdAy__Tst()
Dim O() As Md
O = CurPj.MdAy
Dim I, Md As Md
For Each I In O
    Set Md = I
    Debug.Print Md.Nm
Next
End Sub

Function SrcPth$()
Dim Ffn$: Ffn = Pj.FileName
Dim Fn$: Fn = FfnFn(Ffn)
Dim O$:
O = FfnPth(Pj.FileName) & "Src\": PthEns O
O = O & Fn & "\":                       PthEns O
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
For Each I In MdAy
    Set Md = I
    PushAy O, AyAddPfx(Md.TstFunNy_WithEr, Md.Nm & ".")
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
