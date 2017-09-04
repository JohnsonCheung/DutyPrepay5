Attribute VB_Name = "Vb_Dic"
Option Explicit
Option Compare Database

Function DicAdd(a As Dictionary, B As Dictionary) As Dictionary
Dim O As Dictionary: Set O = DicClone(a)
Dim K
If B.Count > 0 Then
    For Each K In B.Keys
        O.Add K, B(K)
    Next
End If
Set DicAdd = O
End Function

Function DicBrw(a As Dictionary)
DrsBrw DicDrs(a)
End Function

Function DicClone(a As Dictionary) As Dictionary
Dim O As New Dictionary, K
If a.Count > 0 Then
    For Each K In a.Keys
        O.Add K, a(K)
    Next
End If
Set DicClone = O
End Function

Function DicDrs(a As Dictionary) As Drs
Dim O As Drs
O.Fny = SplitSpc("Key Val")
O.Dry = DicDry(a)
DicDrs = O
End Function

Function DicDry(a As Dictionary) As Variant()
Dim O(), I
Dim K(): K = a.Keys
If Not AyIsEmpty(K) Then
    For Each I In K
        Push O, Array(I, a(I))
    Next
End If
DicDry = O
End Function

Function DicVal(a As Dictionary, K) As VOpt
If Not a.Exists(K) Then Exit Function
DicVal = SomV(a(K))
End Function
