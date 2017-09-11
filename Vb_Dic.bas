Attribute VB_Name = "Vb_Dic"
Option Explicit
Option Compare Database
Type KeyVal
    K As String
    V As Variant
End Type
Type KeyValOpt
    Som As Boolean
    KeyVal As KeyVal
End Type

Function DicAdd(A As Dictionary, ParamArray DicAp()) As Dictionary
Dim Av(): Av = DicAp
Dim I, Dic As Dictionary
Dim O As Dictionary
Set O = DicClone(A)
For Each I In Av
    Set Dic = I
    Set O = DicAddOne(O, Dic)
Next
Set DicAdd = O
End Function

Function DicAddKeyPfx(A As Dictionary, Pfx) As Dictionary
Dim O As New Dictionary
If A.Count = 0 Then GoTo X
Dim K
For Each K In A.Keys
    O.Add Pfx & K, A(K)
Next
X:
    Set DicAddKeyPfx = O
End Function

Sub DicAddKeyVal(A As Dictionary, KeyVal As KeyVal)
With KeyVal
    A.Add .K, .V
End With
End Sub

Sub DicAddKeyValOpt(A As Dictionary, KeyValOpt As KeyValOpt)
With KeyValOpt
    If .Som Then DicAddKeyVal A, .KeyVal
End With
End Sub

Function DicAddOne(A As Dictionary, B As Dictionary) As Dictionary
Dim O As Dictionary: Set O = DicClone(A)
Dim K
If B.Count > 0 Then
    For Each K In B.Keys
        O.Add K, B(K)
    Next
End If
Set DicAddOne = O
End Function

Sub DicAssertKey(A As Dictionary, K$)
If Not A.Exists(K) Then Stop
End Sub

Function DicAyAdd(Dy() As Dictionary) As Dictionary
Dim O As Dictionary
    Set O = DicClone(Dy(0))
Dim J%
For J = 1 To UB(Dy)
    Set O = DicAddOne(O, Dy(J))
Next
Set DicAyAdd = O
End Function

Function DicBoolOpt(A As Dictionary, K) As BoolOpt
Dim V As VOpt: V = DicValOpt(A, K)
If V.Som Then DicBoolOpt = SomBool(V.V)
End Function

Function DicBrw(A As Dictionary)
DrsBrw DicDrs(A)
End Function

Function DicClone(A As Dictionary) As Dictionary
Dim O As New Dictionary, K
If A.Count > 0 Then
    For Each K In A.Keys
        O.Add K, A(K)
    Next
End If
Set DicClone = O
End Function

Function DicDrs(A As Dictionary) As Drs
Dim O As Drs
O.Fny = SplitSpc("Key Val")
O.Dry = DicDry(A)
DicDrs = O
End Function

Function DicDry(A As Dictionary) As Variant()
Dim O(), I
If A.Count = 0 Then Exit Function
Dim K(): K = A.Keys
If Not AyIsEmpty(K) Then
    For Each I In K
        Push O, Array(I, A(I))
    Next
End If
DicDry = O
End Function

Function DicDt(A As Dictionary, Optional DtNm$ = "DIc") As Dt
DicDt = Dt(DtNm, DicDrs(A))
End Function

Function DicHasBlankKey(A As Dictionary) As Boolean
If DicIsEmpty(A) Then Exit Function
Dim K
For Each K In A.Keys
    If Trim(K) = "" Then DicHasBlankKey = True: Exit Function
Next
End Function

Function DicIsEmpty(A As Dictionary) As Boolean
DicIsEmpty = A.Count = 0
End Function

Function DicMge(PfxLvs$, ParamArray DicAp()) As Dictionary
Dim Av(): Av = DicAp
Dim Ny$()
    Ny = SplitLvs(PfxLvs)
    Ny = AyAddSfx(Ny, "@")
If Sz(Av) <> Sz(Ny) Then Stop
Dim Dy() As Dictionary
Dim D As Dictionary
    Dim J%
    For J = 0 To UB(Ny)
        Set D = Av(J)
        Push Dy, DicAddKeyPfx(D, Ny(J))
    Next
Set DicMge = DicAyAdd(Dy)
End Function

Function DicMinus(A As Dictionary, B As Dictionary) As Dictionary
If DicIsEmpty(A) Then Set DicMinus = New Dictionary: Exit Function
If DicIsEmpty(B) Then Set DicMinus = DicClone(A): Exit Function
Dim O As New Dictionary, K
For Each K In A.Keys
    If Not B.Exists(K) Then O.Add K, A(K)
Next
Set DicMinus = O
End Function

Function DicVal(A As Dictionary, K, Optional ThrowErIfNotFnd As Boolean)
If Not A.Exists(K) Then
    If ThrowErIfNotFnd Then Stop
    DicVal = "{?}"
    Exit Function
End If
DicVal = A(K)
End Function

Function DicValOpt(A As Dictionary, K) As VOpt
If Not A.Exists(K) Then Exit Function
DicValOpt = SomV(A(K))
End Function

Function KeyVal(K$, V) As KeyVal
KeyVal.K = K
KeyVal.V = V
End Function

Function SomKeyVal(K$, V) As KeyValOpt
SomKeyVal.Som = True
SomKeyVal.KeyVal = KeyVal(K, V)
End Function
