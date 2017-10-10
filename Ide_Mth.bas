Attribute VB_Name = "Ide_Mth"
Option Explicit
Option Compare Database

Function CurMthNm$(Optional A As CodeModule)
Dim Md As CodeModule
    Set Md = DftMd(A)
Dim L&
    Dim R1&, R2&, C1&, C2&
    Md.CodePane.GetSelection R1, C1, R2, C2
    L = R1
Dim K As vbext_ProcKind
CurMthNm = Md.ProcOfLine(L, K)
End Function

Function DftMthNm$(Optional MthNm, Optional A As CodeModule)
If IsEmpty(MthNm) Then
    DftMthNm = CurMthNm(A)
Else
    DftMthNm = MthNm
End If
End Function

Function MthDrsKy(MthDrs As Drs) As String()
Dim Dry() As Variant: Dry = MthDrs.Dry
Dim Fny$(): Fny = MthDrs.Fny
Dim O$()
    Dim Ty$, Mdy$, MthNm$, K$, IdxAy&(), Dr
    IdxAy = FnyLIdxAy(Fny, "Mdy MthNm Ty")
    If AyIsEmpty(MthDrs.Dry) Then Exit Function
    For Each Dr In MthDrs.Dry
        'Debug.Print Mdy, MthNm, Ty
        AyAsg_Idx Dr, IdxAy, Mdy, MthNm, Ty
        Push O, MthKey(Mdy, Ty, MthNm)
    Next
MthDrsKy = O
End Function

Function MthKey$(Mdy$, Ty$, MthNm$)
Dim A1 As Byte
    If IsSfx(MthNm, "__Tst") Then
        A1 = 8
    ElseIf MthNm = "Tst" Then
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
MthKey = FmtQQ("?:?:?", A1, MthNm, A3)
End Function

Sub MthNmAssertIsNotTst(MthNm$)
Const CSub$ = "MthNmAssertIsNotTst"
If IsSfx(MthNm, "__Tst") Then Er CSub, "Given {MthNm} cannot be Sfx-[__Tst]", MthNm
End Sub
