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

Sub DicAssertIsEq(D1 As Dictionary, D2 As Dictionary)
If Not DicIsEq(D1, D2) Then Stop
End Sub

Sub DicAssertKey(A As Dictionary, K$)
If Not A.Exists(K) Then Stop
End Sub

Sub DicAssertKeyLvs(A As Dictionary, KeyLvs$)
DicAssertKy A, SplitLvs(KeyLvs)
End Sub

Sub DicAssertKy(A As Dictionary, Ky)
Dim K
For Each K In Ky
    If Not A.Exists(K) Then Debug.Print K: Stop
Next
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

Function DicAyDr(DicAy, K) As Variant()
Dim U%: U = UB(DicAy)
Dim O()
ReDim O(U + 1)
Dim I, Dic As Dictionary, J%
J = 1
O(0) = K
For Each I In DicAy
    Set Dic = I
    If Dic.Exists(K) Then O(J) = Dic(K)
    J = J + 1
Next
DicAyDr = O
End Function

Function DicAyKy(DicAy) As Variant()
Dim O(), Dic As Dictionary, I
For Each I In DicAy
    Set Dic = I
    PushNoDupAy O, Dic.Keys
Next
DicAyKy = O
End Function

Function DicBoolOpt(A As Dictionary, K) As BoolOpt
Dim V As VOpt: V = DicValOpt(A, K)
If V.Som Then DicBoolOpt = SomBool(V.V)
End Function

Function DicBrw(A As Dictionary)
DrsBrw DicDrs(A)
End Function

Function DicByDry(DicDry) As Dictionary
Dim O As New Dictionary
If Not AyIsEmpty(DicDry) Then
    Dim Dr
    For Each Dr In DicDry
        O.Add Dr(0), Dr(1)
    Next
End If
Set DicByDry = O
End Function

Function DicByFt(Ft) As Dictionary
Set DicByFt = DicByLy(FtLy(Ft))
End Function

Function DicByLines(DicLines$) As Dictionary
Set DicByLines = DicByLy(SplitCrLf(DicLines))
End Function

Function DicByLy(DicLy, Optional IgnoreDup As Boolean) As Dictionary
Const CSub$ = "DicByLy"
Dim O As New Dictionary
    If AyIsEmpty(DicLy) Then Set DicByLy = O: Exit Function
    Dim I
    For Each I In DicLy
        If Trim(I) = "" Then GoTo Nxt
        If FstChr(I) = "#" Then GoTo Nxt
        With Brk(I, " ")
            If O.Exists(.S1) Then
                If Not IgnoreDup Then
                    Er CSub, "Given {DicLy} has duplicate {key}", DicLy, .S1
                End If
            Else
                O.Add .S1, .S2
            End If
        End With
Nxt:
    Next
Set DicByLy = O
End Function

Function DicByStr(DicStr$) As Dictionary
Set DicByStr = DicByLy(SplitVBar(DicStr))
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

Sub DicDmp(A As Dictionary, Optional InclDicValTy As Boolean)
AyDmp DicLy(A, InclDicValTy)
End Sub

Function DicDrs(A As Dictionary, Optional InclDicValTy As Boolean) As Drs
Dim O As Drs
O.Fny = SplitSpc("Key Val"): If InclDicValTy Then Push O.Fny, "ValTy"
O.Dry = DicDry(A, InclDicValTy)
DicDrs = O
End Function

Function DicDry(A As Dictionary, Optional InclDicValTy As Boolean) As Variant()
Dim O(), I
If A.Count = 0 Then Exit Function
Dim K(): K = A.Keys
If Not AyIsEmpty(K) Then
    If InclDicValTy Then
        For Each I In K
            Push O, Array(I, A(I), TypeName(A(I)))
        Next
    Else
        For Each I In K
            Push O, Array(I, A(I))
        Next
    End If
End If
DicDry = O
End Function

Function DicDt(A As Dictionary, Optional DtNm$ = "Dic", Optional InclDicValTy As Boolean) As Dt
DicDt = Dt(DtNm, DicDrs(A, InclDicValTy))
End Function

Function DicEqLy(A As Dictionary) As String()
If A.Count = 0 Then Exit Function
Dim W%
    W = AyWdt(A.Keys)
Dim K, O$()
For Each K In A.Keys
    Push O, K & Space(W - Len(K)) & " = " & A(K)
Next
DicEqLy = O
End Function

Function DicHasBlankKey(A As Dictionary) As Boolean
If DicIsEmpty(A) Then Exit Function
Dim K
For Each K In A.Keys
    If Trim(K) = "" Then DicHasBlankKey = True: Exit Function
Next
End Function

Function DicHasKeyLvs(A As Dictionary, KeyLvs) As Boolean
DicHasKeyLvs = DicHasKy(A, SplitLvs(KeyLvs))
End Function

Function DicHasKy(A As Dictionary, Ky) As Boolean
AssertIsAy Ky
If AyIsEmpty(Ky) Then Stop
Dim K
For Each K In Ky
    If Not A.Exists(K) Then
        Debug.Print FmtQQ("DicHasKy: Key(?) is missing", K)
        Exit Function
    End If
Next
DicHasKy = True
End Function

Function DicIsEmpty(A As Dictionary) As Boolean
If IsNothing(A) Then DicIsEmpty = True: Exit Function
DicIsEmpty = A.Count = 0
End Function

Function DicIsEq(D1 As Dictionary, D2 As Dictionary) As Boolean
If DicIsEmpty(D1) Then Stop
If DicIsEmpty(D2) Then Stop
If D1.Count <> D2.Count Then Exit Function
Dim K1, K2
K1 = AySrt(D1.Keys)
K2 = AySrt(D2.Keys)
If AyIsEq(K1, K2) Then Exit Function
Dim K
For Each K In K1
    If D1(K) <> D2(K) Then Exit Function
Next
DicIsEq = True
End Function

Function DicJn(DicAy, Optional FnyOpt) As Drs
Const CSub$ = "DicJn"
Dim UDic%
    UDic = UB(DicAy)
Dim Fny$()
    If IsEmpty(FnyOpt) Then
        Dim J%
        Push Fny, "Key"
        For J = 0 To UDic
            Push Fny, "V" & J
        Next
    Else
        Fny = FnyOpt
    End If
If UB(Fny) <> UDic + 1 Then Er CSub, "Given {FnyOpt} has {Sz} <> {DicAy-Sz}", FnyOpt, Sz(FnyOpt), Sz(DicAy)
Dim Ky()
    Ky = DicAyKy(DicAy)
Dim URow&
    URow = UB(Ky)
Dim O()
    ReDim O(URow)
    Dim K
    J = 0
    For Each K In Ky
        O(J) = DicAyDr(DicAy, K)
        J = J + 1
    Next
DicJn.Dry = O
DicJn.Fny = Fny
End Function

Function DicKVLy(A As Dictionary) As String()
If DicIsEmpty(A) Then Exit Function
Dim O$(), K, W%, Ky
Ky = A.Keys
W = AyWdt(Ky)
For Each K In Ky
    Push O, AlignL(K, W) & " = " & A(K)
Next
DicKVLy = O
End Function

Function DicLy(A As Dictionary, Optional InclDicValTy As Boolean) As String()
DicLy = DrsLy(DicDrs(A, InclDicValTy))
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

Function DicSelIntoAy(A As Dictionary, Ky$()) As Variant()
Dim O()
Dim U&: U = UB(Ky)
ReDim O(U)
Dim J&
For J = 0 To U
    If Not A.Exists(Ky(J)) Then Stop
    O(J) = A(Ky(J))
Next
DicSelIntoAy = O
End Function

Function DicSrt(A As Dictionary) As Dictionary
If DicIsEmpty(A) Then Set DicSrt = New Dictionary: Exit Function
Dim K
Dim O As New Dictionary
For Each K In AySrt(A.Keys)
    O.Add K, A(K)
Next
Set DicSrt = O
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

Function DicVy(A As Dictionary, Ky) As Variant()
If AyIsEmpty(Ky) Then Exit Function
Dim O()
    Dim K
    For Each K In Ky
        If Not A.Exists(K) Then Stop
        Push O, A(K)
    Next
DicVy = O
End Function

Function DicWs(A As Dictionary) As Worksheet
Set DicWs = DrsWs(DicDrs(A))
End Function

Function DicWsVis(A As Dictionary) As Worksheet
Dim O As Worksheet
    Set O = DicWs(A)
    WsVis O
Set DicWsVis = O
End Function

Function KeyVal(K$, V) As KeyVal
KeyVal.K = K
KeyVal.V = V
End Function

Function SomKeyVal(K$, V) As KeyValOpt
SomKeyVal.Som = True
SomKeyVal.KeyVal = KeyVal(K, V)
End Function
