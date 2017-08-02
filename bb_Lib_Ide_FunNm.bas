Attribute VB_Name = "bb_Lib_Ide_FunNm"
Option Compare Database
Option Explicit
Sub MdFunNmAy__Tst()
BrwAy MdFunNmAy
End Sub
Sub PjFunNmAy__Tst()
BrwAy PjFunNmAy
End Sub
Sub AAA()
MdFunNmAy__Tst
End Sub
Function PjFunNmAy(Optional A As VBProject) As String()
End Function
Function MdFunNmAy(Optional A As CodeModule) As String()
Dim O$()
    Dim L$, Lin
    For Each Lin In MdLy(JnContLin:=True, A:=A)
        L = SrcLinFunNm(Lin)
        If L <> "" Then Push O, L
    Next
MdFunNmAy = O
End Function
Sub SrcLinFunNm__Tst()
Dim Act$: Act = SrcLinFunNm("Public Function AA()")
Debug.Assert Act = "AA:Public:Function:"
End Sub
Function SrcLinFunNm$(SrcLin)
Dim O$: O = SrcLin
Dim Mdy$: Mdy = ParseMdy(O)
Dim FunTy$: FunTy = ParseFunTy(O): FunTy = RmvPFx(FunTy, "Property ")
If FunTy = "" Then Exit Function
Dim Nm$: Nm = ParseNm(O)
SrcLinFunNm = Nm & ":" & Mdy & ":" & FunTy
End Function
Function ParseMdy$(OLin$)
ParseMdy = ParseOneOf(OLin, Sy("Public", "Private ", "Friend"))
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

