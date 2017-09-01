Attribute VB_Name = "Vb_Dic"
Option Compare Database
Option Explicit

Function DicAdd(A As Dictionary, B As Dictionary) As Dictionary
Dim O As Dictionary: Set O = DicClone(A)
Dim K
If B.Count > 0 Then
    For Each K In B.Keys
        O.Add K, B(K)
    Next
End If
Set DicAdd = O
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
Dim K(): K = A.Keys
If Not AyIsEmpty(K) Then
    For Each I In K
        Push O, Array(I, A(I))
    Next
End If
DicDry = O
End Function

Function DicVal(A As Dictionary, K) As VOpt
If Not A.Exists(K) Then Exit Function
DicVal = SomV(A(K))
End Function
