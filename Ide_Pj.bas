Attribute VB_Name = "Ide_Pj"
Option Explicit
Option Compare Database

Function DftPj(Optional A As Vbproject) As Vbproject
If IsNothing(A) Then
    Set DftPj = Application.Vbe.ActiveVBProject
Else
    Set DftPj = A
End If
End Function

Function Pj(PjNm) As Vbproject
Set Pj = Application.Vbe.VBProjects(PjNm)
End Function

Sub PjAssertNotUnderSrc(Optional A As Vbproject)
Dim B$: B = PjPth(A)
If PthFdr(B) = "Src" Then Stop
End Sub

Sub PjCpyToSrc(Optional A As Vbproject)
FilCpyToPth DftPj(A).FileName, PjSrcPth(A), OvrWrt:=True
End Sub

Sub PjEnsOptExplicit(Optional A As Vbproject)
Dim I, Md As CodeModule
For Each I In PjMdAy(, A)
    Set Md = I
    MdEnsOptExplicit Md
Next
End Sub

Sub PjExp(Optional A As Vbproject)
PjAssertNotUnderSrc
PjCpyToSrc A
PthClrFil PjSrcPth(A)
Dim Md As CodeModule, I
For Each I In PjMdAy(, A)
    Set Md = I
    MdExp Md
Next
End Sub

Function PjHasMd(MdNm, Optional A As Vbproject) As Boolean
Dim Cmp As VBComponent
For Each Cmp In DftPj(A).VBComponents
    If MdNm = Cmp.Name Then
        PjHasMd = True
        Exit Function
    End If
Next
End Function

Function PjMdAy(Optional LikMdNm$ = "*", Optional A As Vbproject) As CodeModule()
Dim O() As CodeModule
Dim Cmp As VBComponent
For Each Cmp In DftPj(A).VBComponents
    If Cmp.Name Like LikMdNm Then
        PushObj O, Cmp.CodeModule
    End If
Next
PjMdAy = O
End Function

Function PjMdNy(Optional LikMdNm$ = "*", Optional A As Vbproject) As String()
PjMdNy = OyPrp(PjMdAy(LikMdNm, A), "Name", EmptySy)
End Function

Function PjMthDotNy(Optional MthNmLik = "*", Optional A As Vbproject)
Dim O$(), I, Md As CodeModule, Ay$()
For Each I In PjMdAy(, A)
    Set Md = I
    PushAy O, AyAddPfx(MdMthNy(MthNmLik, Md), MdNm(Md) & ".")
Next
PjMthDotNy = O
End Function

Function PjMthDrs(Optional WithBdyLy As Boolean, Optional WithBdyLines As Boolean, Optional A As Vbproject) As Drs
Dim Fny$()
    Fny = SrcMthDrsFny(WithBdyLy, WithBdyLines)
    Push Fny, "MdNm"
PjMthDrs.Fny = Fny
PjMthDrs.Dry = PjMthDry(WithBdyLy, WithBdyLines, A)
End Function

Function PjMthDry(Optional WithBdyLy As Boolean, Optional WithBdyLines As Boolean, Optional A As Vbproject) As Variant()
Dim Dry()
    Dim I, Md As CodeModule
    For Each I In PjMdAy(, A)
        Set Md = I
        PushAy Dry, MdMthDrs(WithBdyLy, WithBdyLines, A:=Md).Dry
    Next
PjMthDry = Dry
End Function

Function PjNm$(Optional A As Vbproject)
PjNm = DftPj(A).Name
End Function

Function PjPth$(Optional A As Vbproject)
PjPth = FfnPth(DftPj(A).FileName)
End Function

Sub PjSrcBrw()
PthBrw PjSrcPth
End Sub

Function PjSrcPth$(Optional A As Vbproject)
Dim Ffn$: Ffn = DftPj(A).FileName
Dim Fn$: Fn = FfnFn(Ffn)
Dim O$:
O = FfnPth(DftPj(A).FileName) & "Src\": PthEns O
O = O & Fn & "\":                       PthEns O
PjSrcPth = O
End Function

Sub PjSrt(Optional A As Vbproject)
Dim Md As CodeModule, I
For Each I In PjMdAy(, A)
    Set Md = I
    If MdNm(Md) <> "Ide" Then
        MdSrt Md
    End If
Next
End Sub

Sub PjSrtRpt(Optional A As Vbproject)
Dim Md As CodeModule, I
For Each I In PjMdAy(, A)
    Set Md = I
    MdSrtRpt Md
Next
End Sub

Sub PjSrtRptDif(Optional A As Vbproject)
Dim Md As CodeModule, I
For Each I In PjMdAy(, A)
    Set Md = I
    MdSrtRptDif Md
Next
End Sub

Function PjTstMthNy_WithEr(Optional A As Vbproject) As String()
Dim O$(), I, Md As CodeModule
For Each I In PjMdAy(, A)
    Set Md = I
    PushAy O, AyAddPfx(MdTstMthNy_WithEr(Md), MdNm(Md) & ".")
Next
PjTstMthNy_WithEr = O
End Function

Sub PjUpdTstMth(Optional A As Vbproject)
Dim I, Md As CodeModule
For Each I In PjMdAy(, A)
    Set Md = I
    MdUpdTstMth Md
Next
End Sub

Private Sub PjMdAy__Tst()
Dim O() As CodeModule
O = PjMdAy
Dim I, Md As CodeModule
For Each I In O
    Set Md = I
    Debug.Print MdNm(Md)
Next
End Sub

Private Sub PjMdNy__Tst()
AyBrw PjMdNy
End Sub

Private Sub PjMthDrs__Tst()
Dim Drs As Drs
Drs = PjMthDrs(WithBdyLines:=True)
WsVis DrsWs(Drs, PjNm)
End Sub
