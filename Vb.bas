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

Function CollObjAy(ObjColl) As Object()
Dim O() As Object
Dim V
For Each V In ObjColl
    Push O, V
Next
CollObjAy = O
End Function

Function Dft(V, DftVal)
If IsEmpty(V) Then
    Dft = DftVal
Else
    Dft = V
End If
End Function

Sub Er(MacroStr$, ParamArray Ap())
Dim Av(): Av = Ap
AyBrw MsgLy(MacroStr, Av)
Stop
End Sub

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

Function MsgAyLy(MsgAy()) As String()
Dim I, Av(), O$(), MacroStr$
For Each I In MsgAy
    Av = I
    MacroStr = AyShift(Av)
    PushAy O, MsgLy(MacroStr, Av)
Next
MsgAyLy = O
End Function

Sub MsgBrw(MacroStr$, Av())
AyBrw MsgLy(MacroStr, Av())
End Sub

Function MsgLy(MacroStr$, Av()) As String()
Dim Ny$(): Ny = MacroStrNy(MacroStr)
Dim O$()
    PushAy O, SplitVBar(MacroStr)
Dim I, J%
For Each I In Ny
    Push O, Chr(9) & I
    PushAy O, AyAddPfx(VarLy(Av(J)), Chr(9) & Chr(9))
Next
MsgLy = O
End Function

Property Get NowStr$()
NowStr = Format(Now(), "YYYY-MM-DD HH:MM:SS")
End Property

Function PipeAy(Prm, FunNy$())
Dim O: Asg Prm, O
Dim I
For Each I In FunNy
    Asg Run(I, O), O
Next
Asg O, PipeAy
End Function

Function RestTerm$(S)
RestTerm = Brk1(Trim(S), " ").S2
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

