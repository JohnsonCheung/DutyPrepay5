Attribute VB_Name = "Vb_Opt"
Option Explicit
Option Compare Database
Type StrOpt
    Som As Boolean
    Str As String
End Type
Type BoolOpt
    Bool As Boolean
    Som As Boolean
End Type
Type SyOpt
    Som As Boolean
    Sy() As String
End Type
Type VOpt
    Som As Boolean
    V As Variant
End Type
Type BoolAyOpt
    BoolAy() As Boolean
    Som As Boolean
End Type
Type S1S2Opt
    Som As Boolean
    S1S2 As S1S2
End Type

Function BoolAy_And(A() As Boolean) As Boolean
Dim I
For Each I In A
    If Not I Then Exit Function
Next
BoolAy_And = True
End Function

Function BoolAy_Or(A() As Boolean) As Boolean
Dim I
If AyIsEmpty(A) Then Exit Function
For Each I In A
    If I Then BoolAy_Or = True: Exit Function
Next
End Function

Function BoolAyOpt_And(A As BoolAyOpt) As BoolOpt
If Not A.Som Then Exit Function
BoolAyOpt_And = SomBool(BoolAy_And(A.BoolAy))
End Function

Function BoolAyOpt_Or(A As BoolAyOpt) As BoolOpt
If Not A.Som Then Exit Function
BoolAyOpt_Or = SomBool(False)
End Function

Function SomBool(Bool) As BoolOpt
SomBool.Som = True
SomBool.Bool = Bool
End Function

Function SomBoolAy(A() As Boolean) As BoolAyOpt
SomBoolAy.Som = True
SomBoolAy.BoolAy = A
End Function

Function SomS1S2(A As S1S2) As S1S2Opt
SomS1S2.S1S2 = A
SomS1S2.Som = True
End Function

Function SomStr(S) As StrOpt
SomStr.Som = True
SomStr.Str = S
End Function

Function SomSy(Sy$()) As SyOpt
SomSy.Som = True
SomSy.Sy = Sy
End Function

Function SomV(V) As VOpt
SomV.Som = True
SomV.V = V
End Function
