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
Function IsEmpty(V) As Boolean
IsEmpty = True
If IsMissing(V) Then Exit Function
If IsNothing(V) Then Exit Function
If IsStr(V) Then
    If V = "" Then Exit Function
End If
If IsArray(V) Then
    If IsEmptyAy(V) Then Exit Function
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
