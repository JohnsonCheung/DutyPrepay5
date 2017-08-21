VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Function Md(Optional MdNm$) As Md
Dim M As CodeModule
If MdNm = "" Then
    Set M = CurMd
Else
    Set M = CurPj.VBComponents(MdNm).CodeModule
End If
Set Md = Nw.Md(M)
End Function

