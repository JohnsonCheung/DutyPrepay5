Attribute VB_Name = "Tst"
Option Explicit
Option Compare Database
Type ThowMsgOrStr
    Som As Boolean
    Str As String
    ThowMsg As String
End Type
Type ThowMsgOrSy
    Som As Boolean
    Sy() As String
    ThowMsg As String
End Type
Type ThowMsgOrInt
    Som As Boolean
    Int As Integer
    ThowMsg As String
End Type
Type ThowMsgOrVar
    Som As Boolean
    V As Variant
    ThowMsg As String
End Type

Sub AssertActEqExp(Act, Exp)
If VarType(Act) <> VarType(Exp) Then Stop
If IsPrim(Act) Then
    If Act <> Exp Then
        Debug.Print "Act:"; Act; "<"
        Debug.Print "Exp:"; Exp; "<"
        Stop
    End If
    Exit Sub
End If
If IsArray(Act) Then
    If Not AyIsEq(Act, Exp) Then Stop
    Exit Sub
End If
Stop
End Sub

Function TMOIntSomInt(I%) As ThowMsgOrInt
TMOIntSomInt.Som = True
TMOIntSomInt.Int = I
End Function

Function TMOIntThowMsg(ThowMsg$) As ThowMsgOrInt
TMOIntThowMsg.ThowMsg = ThowMsg
End Function

Function TMOStrDmp(A As ThowMsgOrStr, Optional Nm$ = "ThowMsgOrStr")
With A
    Debug.Print Nm$; " = ";
    Debug.Print IIf(.Som, "SomStr ", "SomThowMsg ");
    Debug.Print IIf(.Som, .Str, .ThowMsg)
End With
End Function

Function TMOStrSomStr(Str$) As ThowMsgOrStr
TMOStrSomStr.Som = True
TMOStrSomStr.Str = Str
End Function

Function TMOStrThowMsg(ThowMsg$) As ThowMsgOrStr
TMOStrThowMsg.ThowMsg = ThowMsg
End Function

Function TMOSySomSy(Sy$()) As ThowMsgOrSy
TMOSySomSy.Som = True
TMOSySomSy.Sy = Sy
End Function

Function TMOSySomThowMsg(ThowMsg$) As ThowMsgOrSy
TMOSySomThowMsg.ThowMsg = ThowMsg
End Function

Function TstResFdr$(Fdr$)
Dim O$
    O = TstResPth & Fdr & "\"
    PthEns O
TstResFdr = O
End Function

Sub TstResFdrBrw(Fdr$)
PthBrw TstResFdr(Fdr)
End Sub

Function TstResPth$()
Dim O$
    O = PjSrcPth & "TstRes\"
    PthEns O
TstResPth = O
End Function

Sub TstResPthBrw()
PthBrw TstResPth
End Sub
