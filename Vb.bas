Attribute VB_Name = "Vb"
Option Explicit
Option Compare Database
Public Fso As New FileSystemObject

Sub Asg(V, OV)
If IsObject(V) Then
    Set OV = V
Else
    OV = V
End If
End Sub

Sub AssertIsPfx(S, I)
If Not IsPfx(S, I) Then Er "AssertIsPfx", "{S} does not have this {Pfx}", S, I
End Sub
Sub StrOptDmp(A As StrOpt)
Debug.Print StrOptToStr(A)
End Sub
Function StrOptToStr(A As StrOpt)
StrOptToStr = "StrOpt: *Som=" & A.Som & vbCrLf & A.Str
End Function
Sub AssertIsStr(V)
If Not IsStr(V) Then Stop
End Sub

Sub AssertIsSy(V)
If Not IsSy(V) Then Er "AssertIsSy", "VarType-{V} is not String Array", TypeName(V)
End Sub

Sub AssertNonEmpty(V)
If IsEmpty(V) Then Er "AssertNonEmpty", "Given V is IsEmpty"
End Sub

Sub AssertNotHasSubStr(S, SubStr)
AssertIsStr S
AssertIsStr SubStr
If HasSubStr(S, SubStr) Then Er "AssertNotHasSubStr", "{S} cannot has {SubStr}", S, SubStr
End Sub

Function CollObjAy(ObjColl) As Object()
Dim O() As Object
Dim V
For Each V In ObjColl
    Push O, V
Next
CollObjAy = O
End Function

Function Dft(V, DftV)
If IsEmpty(V) Then
    Dft = V
    Dft = DftV
End If
End Function

Sub Er(CSub$, MacroStr$, ParamArray Ap())
Dim Av(): Av = Ap
AyBrw ErMsgLy(CSub, MacroStr, Av)
Stop
End Sub

Sub ErMsgBrw(CSub$, MacroStr$, Av())
AyBrw ErMsgLy(CSub, MacroStr, Av())
End Sub

Function ErMsgLy(CSub$, MacroStr$, Av()) As String()
Dim O$()
    Push O, "Subr-" & CSub & ": " & RplVbar(MacroStr)
If Not AyIsEmpty(Av) Then
    Dim Ny$(): Ny = MacroStrNy(MacroStr)
    Dim I, J%
    If Not AyIsEmpty(Ny) Then
        For Each I In Ny
            Push O, Chr(9) & I
            PushAy O, AyAddPfx(VarLy(Av(J)), Chr(9) & Chr(9))
            J = J + 1
        Next
    End If
End If
ErMsgLy = O
End Function

Function FstTerm$(S)
FstTerm = Brk1(Trim(S), " ").S1
End Function

Function IsBool(V) As Boolean
IsBool = VarType(V) = vbBoolean
End Function

Function IsEmpty(V) As Boolean
IsEmpty = True
If IsMissing(V) Then Exit Function
If IsNothing(V) Then Exit Function
If VBA.IsEmpty(V) Then Exit Function
If IsStr(V) Then
    If V = "" Then Exit Function
End If
If IsArray(V) Then
    If AyIsEmpty(V) Then Exit Function
End If
IsEmpty = False
End Function

Function IsEmptyColl(ObjColl) As Boolean
IsEmptyColl = (ObjColl.Count = 0)
End Function

Function IsIn(V, ParamArray Ap()) As Boolean
Dim Av(): Av = Ap
Dim I
For Each I In Av
    If I = V Then IsIn = True: Exit Function
Next
End Function

Function IsIntAy(V) As Boolean
IsIntAy = VarType(V) = vbArray + vbInteger
End Function

Function IsLngAy(V) As Boolean
IsLngAy = VarType(V) = vbArray + vbLong
End Function

Function IsNothing(V) As Boolean
IsNothing = TypeName(V) = "Nothing"
End Function

Function IsPrim(V) As Boolean
Select Case VarType(V)
Case _
    VbVarType.vbBoolean, _
    VbVarType.vbByte, _
    VbVarType.vbCurrency, _
    VbVarType.vbDate, _
    VbVarType.vbDecimal, _
    VbVarType.vbDouble, _
    VbVarType.vbInteger, _
    VbVarType.vbLong, _
    VbVarType.vbSingle, _
    VbVarType.vbString
    IsPrim = True
End Select
End Function

Function IsStr(V) As Boolean
IsStr = VarType(V) = vbString
End Function

Function IsStrAy(V) As Boolean
IsStrAy = VarType(V) = vbArray + vbString
End Function

Function IsSy(V) As Boolean
IsSy = IsStrAy(V)
End Function

Function JnVBar$(Ay)
JnVBar = Join(Ay, "|")
End Function

Function Max(ParamArray Ap())
Dim Av(), O
Av = Ap
O = Av(0)
Dim J%
For J = 1 To UB(Av)
    If Av(J) > O Then O = Av(J)
Next
Max = O
End Function

Function Min(ParamArray Ap())
Dim Av(), O
Av = Ap
O = Av(0)
Dim J%
For J = 1 To UB(Av)
    If Av(J) < O Then O = Av(J)
Next
Min = O
End Function

Sub Never()
Const CSub$ = "Never"
Er CSub, "Should never reach here"
End Sub

Property Get NowStr$()
NowStr = Format(Now(), "YYYY-MM-DD HH:MM:SS")
End Property

Function ObjAyPrpSy(ObjAy, PrpNm$) As String()
If AyIsEmpty(ObjAy) Then Exit Function
Dim O$(), I
For Each I In ObjAy
    Push O, ObjPrp(I, PrpNm)
Next
ObjAyPrpSy = O
End Function

Function ObjAySelPrp(Oy, PrpNm$, PrpVal)
Dim O
    O = Oy
    Erase O
If Not AyIsEmpty(Oy) Then
    Dim I
    For Each I In Oy
        If CallByName(I, PrpNm, VbGet) = PrpVal Then PushObj O, I
    Next
End If
ObjAySelPrp = O
End Function

Function ObjPrp(Obj, PrpNm$)
ObjPrp = CallByName(Obj, PrpNm, VbGet)
End Function

Function PipeAy(Prm, MthNy$())
Dim O: Asg Prm, O
Dim I
For Each I In MthNy
    Asg Run(I, O), O
Next
Asg O, PipeAy
End Function

Function RmvFstChr$(S)
RmvFstChr = Mid(S, 2)
End Function

Function RplFstChr$(S, To_Str)
RplFstChr = To_Str & RmvFstChr(S)
End Function

Function RplFstChrToIf$(S, To_Str, If_Str)
If FstChr(S) = If_Str Then
    RplFstChrToIf = RplFstChr(S, To_Str)
Else
    RplFstChrToIf = S
End If
End Function

Function RstTerm$(S)
RstTerm = Brk1(Trim(S), " ").S2
End Function

Function VarLy(V) As String()
If IsPrim(V) Then
    VarLy = Sy(V)
ElseIf IsArray(V) Then
    VarLy = AySy(V)
ElseIf IsObject(V) Then
    VarLy = Sy("*Type: " & TypeName(V))
Else
    Stop
End If
End Function

Private Sub IsStrAy__Tst()
Dim A$()
Dim B: B = A
Dim C()
Dim D
Debug.Assert IsStrAy(A) = True
Debug.Assert IsStrAy(B) = True
Debug.Assert IsStrAy(C) = False
Debug.Assert IsStrAy(D) = False
End Sub

Sub Tst()
IsStrAy__Tst
End Sub

