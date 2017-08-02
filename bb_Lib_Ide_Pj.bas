Attribute VB_Name = "bb_Lib_Ide_Pj"
Option Compare Database
Option Explicit
Sub ExpPj(Optional A As VBProject)
ClrPthFil PjSrcPth(A)
Dim Md As CodeModule, I
For Each I In PjMdAy(A)
    Set Md = I
    ExpMd Md
Next
End Sub
Function DftPj(Optional A As VBProject) As VBProject
If IsNothing(A) Then
    Set DftPj = Application.VBE.ActiveVBProject
Else
    Set DftPj = A
End If
End Function
Sub PjMdAy__Tst()
Dim O() As CodeModule
O = PjMdAy
Dim I, Md As CodeModule
For Each I In O
    Set Md = I
    Debug.Print MdNm(Md)
Next
End Sub
Function PjMdAy(Optional A As VBProject) As CodeModule()
Dim O() As CodeModule
Dim Cmp As VBComponent
For Each Cmp In DftPj(A).VBComponents
    PushObj O, Cmp.CodeModule
Next
PjMdAy = O
End Function
Sub BrwPjSrc()
BrwPth PjSrcPth
End Sub
Function PjSrcPth$(Optional A As VBProject)
Dim Ffn$: Ffn = DftPj(A).FileName
Dim Fn$: Fn = FfnFn(Ffn)
Dim O$: O = FfnPth(DftPj(A).FileName) & "Src\" & Fn & "\"
EnsPth O
PjSrcPth = O
End Function
Function MdSrcExt$(Optional A As CodeModule)
Dim O$
Select Case MdCmpTy(A)
Case vbext_ct_ClassModule: O = ".cls"
Case vbext_ct_Document: O = ".cls"
Case vbext_ct_StdModule: O = ".bas"
Case Else: Err.Raise 1, , "MdSrcExt: Unexpected MdCmpTy.  Should be [Class or Module or Document]"
End Select
MdSrcExt = O
End Function
Function MdCmpTy(Optional A As CodeModule) As vbext_ComponentType
MdCmpTy = MdCmp(A).Type
End Function
Function MdCmp(Optional A As CodeModule) As VBComponent
Set MdCmp = DftMd(A).Parent
End Function

Function MdSrcFn$(Optional A As CodeModule)
MdSrcFn = MdCmp(A).Name & MdSrcExt(A)
End Function
Function MdSrcFfn$(Optional A As CodeModule)
MdSrcFfn = PjSrcPth(MdPj(A)) & MdSrcFn(A)
End Function
Function MdPj(Optional A As CodeModule) As VBProject
Set MdPj = DftMd(A).Parent.Collection.Parent
End Function
Function ExpMd(Optional A As CodeModule)
Dim Md As CodeModule: Set Md = DftMd(A)
Md.Parent.Export MdSrcFfn(Md)
Debug.Print MdNm(A)
End Function
Function MdNm(Optional A As CodeModule)
MdNm = DftMd(A).Parent.Name
End Function
Function DftMd(A As CodeModule) As CodeModule
If IsNothing(A) Then
    Set DftMd = Application.VBE.ActiveCodePane.CodeModule
Else
    Set DftMd = A
End If
End Function
