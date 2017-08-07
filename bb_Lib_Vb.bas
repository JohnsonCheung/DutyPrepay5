Attribute VB_Name = "bb_Lib_Vb"
Option Compare Database
Option Explicit
Public Fso As New FileSystemObject
Property Get NowStr$()
NowStr = Format(Now(), "YYYY-MM-DD HH:MM:SS")
End Property
Function IsEmptyColl(ObjColl) As Boolean
IsEmptyColl = (ObjColl.Count = 0)
End Function
Function CollObjAy(ObjColl) As Object()
Dim O() As Object
Dim V
For Each V In ObjColl
    Push O, V
Next
CollObjAy = O
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
Function IsStrAy(V) As Boolean
IsStrAy = VarType(V) = vbArray + vbString
End Function
Function IsEmpty(V) As Boolean
IsEmpty = True
If IsMissing(V) Then Exit Function
If IsNothing(V) Then Exit Function
If IsStr(V) Then
    If V = "" Then Exit Function
End If
If IsArray(V) Then
    If AyIsEmpty(V) Then Exit Function
End If
IsEmpty = False
End Function
Function IsStr(V) As Boolean
IsStr = VarType(V) = vbString
End Function
Function Dft(V, DftVal)
If IsEmpty(V) Then
    Dft = DftVal
Else
    Dft = V
End If
End Function
Sub Tst()
IsStrAy__Tst
End Sub
