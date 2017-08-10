Attribute VB_Name = "bb_Lib_Vb_Opt"
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
Function SomSy(Sy$()) As SyOpt
SomSy.Som = True
SomSy.Sy = Sy
End Function
Function SomBool(Bool) As BoolOpt
SomBool.Som = True
SomBool.Bool = Bool
End Function
Function SomStr(S) As StrOpt
SomStr.Som = True
SomStr.Str = S
End Function

