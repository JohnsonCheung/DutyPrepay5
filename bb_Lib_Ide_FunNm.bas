Attribute VB_Name = "bb_Lib_Ide_FunNm"
Option Compare Database
Option Explicit
Sub MdFunDrs__Tst()
BrwDrs MdFunDrs
End Sub
Sub PjFunDrs__Tst()
WsVis DtWs(PjFunDrs)
End Sub
Sub AAA()
PjFunDrs__Tst
End Sub
Function PjFunDrs(Optional A As VBProject) As Dt
Dim Dry()
    Dim I, Md As CodeModule
    For Each I In PjMdAy(A)
        Set Md = I
        PushAy Dry, MdFunDrs(Md).Dry
    Next
Dim O As Dt
    O.Fny = SplitSpc("Mdy Ty FunNm MdNm")
    O.Dry = Dry
PjFunDrs = O
End Function
Function IsEmptyMd(Optional A As CodeModule) As Boolean
IsEmptyMd = DftMd(A).CountOfLines = 0
End Function
Function MdFunDrs(Optional A As CodeModule) As Drs
Dim Dry()
    If Not IsEmptyMd(A) Then
        Dim Dr(), Lin
        Dim Nm$: Nm = MdNm(A)
        For Each Lin In MdLy(JnContLin:=True, A:=A)
            Dr = SrcLinFunDr(Lin)
            If Not IsEmptyAy(Dr) Then
                Push Dr, Nm
                Push Dry, Dr
            End If
        Next
    End If
Dim O As Drs
    O.Fny = SplitSpc("Mdy Ty FunNm MdNm")
    O.Dry = Dry
MdFunDrs = O
End Function
Sub SrcLinFunDr__Tst()
Dim Act(): Act = SrcLinFunDr("Private Function AA()")
Debug.Assert Sz(Act) = 3
Debug.Assert Act(0) = "Private"
Debug.Assert Act(1) = "Function"
Debug.Assert Act(2) = "AA"
End Sub
Function SrcLinFunDr(SrcLin) As Variant()
Dim O$: O = SrcLin
Dim Mdy$: Mdy = ParseMdy(O)
Dim Ty$: Ty = ParseFunTy(O): Ty = RmvPFx(Ty, "Property ")
If Ty = "" Then Exit Function
Dim Nm$: Nm = ParseNm(O)
SrcLinFunDr = Array(Mdy, Ty, Nm)
End Function
Function ParseMdy$(OLin$)
ParseMdy = ParseOneOf(OLin, Sy("Public", "Private", "Friend"))
End Function
Function ParseFunTy$(OLin$)
ParseFunTy = ParseOneOf(OLin, Sy("Function", "Sub", "Property Get", "Property Let", "Property Set", "Type", "Enum"))
End Function
Function ParseOneOf(OLin$, OneOfAy$())
Dim I
For Each I In OneOfAy
    If IsPfx(OLin, I) Then OLin = RmvFstNChr(RmvPFx(OLin, I)): ParseOneOf = I: Exit Function
Next
End Function
Private Function IsPfx(S, Pfx) As Boolean
IsPfx = Left(S, Len(Pfx)) = Pfx
End Function

Function ParseNm$(OLin$)
Dim J%
J = 1
If Not IsLetter(FstChr(OLin)) Then GoTo Nxt
For J = 2 To Len(OLin)
    If Not IsNmChr(Mid(OLin, J, 1)) Then GoTo Nxt
Next
Nxt:
If J = 1 Then Exit Function
ParseNm = Left(OLin, J - 1)
OLin = Mid(OLin, J)
End Function
Function IsNmChr(C) As Boolean
IsNmChr = True
If IsLetter(C) Then Exit Function
If C = "_" Then Exit Function
If IsDigit(C) Then Exit Function
IsNmChr = False
End Function
Function IsDigit(C) As Boolean
IsDigit = "0" <= C And C <= "9"
End Function
Function IsLetter(C) As Boolean
Dim C1$: C1 = UCase(C)
IsLetter = ("A" <= C1 And C1 <= "Z")
End Function

