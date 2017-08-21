VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Md_Src_Static"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Function FunKey$(Mdy$, Ty$, FunNm$)
Dim A1 As Byte
    If IsSfx(FunNm, "__Tst") Then
        A1 = 8
    ElseIf FunNm = "Tst" Then
        A1 = 9
    Else
        Select Case Mdy
        Case "Public", "": A1 = 1
        Case "Friend": A1 = 2
        Case "Private": A1 = 3
        Case Else: Stop
        End Select
    End If
Dim A3$
    If Ty <> "Function" And Ty <> "Sub" Then A3 = Ty
FunKey = FmtQQ("?:?:?", A1, FunNm, A3)
End Function
