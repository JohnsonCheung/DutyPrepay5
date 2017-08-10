Attribute VB_Name = "bb_Lib_Vb"
Option Compare Database
Option Explicit
Public Fso As New FileSystemObject

Function CollObjAy(ObjColl) As Object()
Dim O() As Object
Dim V
For Each V In ObjColl
    Push O, V
Next
CollObjAy = O
End Function
Function FstTerm$(S)
FstTerm = Brk1(Trim(S), " ").S1
End Function
Function RestTerm$(S)
RestTerm = Brk1(Trim(S), " ").S2
End Function
Function Dft(V, DftVal)
If IsEmpty(V) Then
    Dft = DftVal
Else
    Dft = V
End If
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

Function IsStr(V) As Boolean
IsStr = VarType(V) = vbString
End Function

Function IsStrAy(V) As Boolean
IsStrAy = VarType(V) = vbArray + vbString
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

Property Get NowStr$()
NowStr = Format(Now(), "YYYY-MM-DD HH:MM:SS")
End Property

Function VarLen%(V)
If IsNull(V) Then Exit Function
If IsArray(V) Then
    If AyIsEmpty(V) Then Exit Function
    VarLen = Len(V(0))
    Exit Function
End If
VarLen = Len(V)
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
