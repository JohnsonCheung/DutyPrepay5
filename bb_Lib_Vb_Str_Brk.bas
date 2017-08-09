Attribute VB_Name = "bb_Lib_Vb_Str_Brk"
Option Compare Database
Option Explicit
Type S1S2
    S1 As String
    S2 As String
End Type
Type Map
    Sy1() As String
    Sy2() As String
End Type

Function Brk(S, Sep, Optional NoTrim As Boolean) As S1S2
Dim P&: P = InStr(S, Sep)
If P = 0 Then Err.Raise "Brk: Str[" & S & "] does not contains Sep[" & Sep & "]"
Brk = BrkAt(S, P, Len(Sep), NoTrim)
End Function

Function Brk1(S, Sep, Optional NoTrim As Boolean) As S1S2
Dim P&: P = InStr(S, Sep)
If P = 0 Then
    Dim O As S1S2
    If NoTrim Then
        O.S1 = S
    Else
        O.S1 = Trim(S)
    End If
    Brk1 = O
    Exit Function
End If
Brk1 = BrkAt(S, P, Len(Sep), NoTrim)
End Function

Function Brk2(S, Sep, Optional NoTrim As Boolean) As S1S2
Dim P&: P = InStr(S, Sep)
If P = 0 Then
    Dim O As S1S2
    If NoTrim Then
        O.S2 = S
    Else
        O.S2 = Trim(S)
    End If
    Brk2 = O
    Exit Function
End If
Brk2 = BrkAt(S, P, Len(Sep), NoTrim)
End Function

Function BrkAt(S, P&, SepLen%, Optional NoTrim As Boolean) As S1S2
Dim O As S1S2
With O
    If NoTrim Then
        .S1 = Left(S, P - 1)
        .S2 = Mid(S, P + SepLen)
    Else
        .S1 = Trim(Left(S, P - 1))
        .S2 = Trim(Mid(S, P + SepLen))
    End If
End With
BrkAt = O
End Function

Function BrkBoth(S, Sep, Optional NoTrim As Boolean) As S1S2
Dim P&: P = InStr(S, Sep)
If P = 0 Then
    Dim O As S1S2
    If NoTrim Then
        O.S1 = S
    Else
        O.S1 = Trim(S)
    End If
    O.S2 = O.S1
    BrkBoth = O
    Exit Function
End If
BrkBoth = BrkAt(S, P, Len(Sep), NoTrim)
End Function

Function BrkMapStr(MapStr$) As Map
Dim Ay$(): Ay = Split(MapStr, "|")
Dim Ay1$(), Ay2$()
    Dim I
    For Each I In Ay
        With BrkBoth(I, ":")
            Push Ay1, .S1
            Push Ay2, .S2
        End With
    Next
Dim O As Map
    O.Sy1 = Ay1
    O.Sy2 = Ay2
BrkMapStr = O
End Function

Function BrkQuote(QuoteStr$) As S1S2
Dim L%: L = Len(QuoteStr)
Dim O As S1S2
Select Case L
Case 0:
Case 1
    O.S1 = QuoteStr
    O.S2 = O.S1
Case 2
    O.S1 = Left(QuoteStr, 1)
    O.S2 = Right(QuoteStr, 1)
Case Else
    Dim P%
    If InStr(QuoteStr, "*") > 0 Then
        O = Brk(QuoteStr, "*", NoTrim:=True)
    End If
End Select
BrkQuote = O
End Function

Private Sub BrkMapStr__Tst()
Dim MapStr$
MapStr = "aa:bb|cc|dd:ee"
Dim Act As Map: Act = BrkMapStr(MapStr)
Dim Exp1$(): Exp1 = SplitSpc("aa cc dd"): AssertEqAy Exp1, Act.Sy1
Dim Exp2$(): Exp2 = SplitSpc("aa cc dd"): AssertEqAy Exp2, Act.Sy1
End Sub

Sub Tst()
BrkMapStr__Tst
End Sub
