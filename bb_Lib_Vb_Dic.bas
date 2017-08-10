Attribute VB_Name = "bb_Lib_Vb_Dic"
Option Compare Database
Option Explicit
Function DicBrw(A As Dictionary)
DrsBrw DicDrs(A)
End Function
Function DicDrs(A As Dictionary) As Drs
Dim Dry(), I
Dim K(): K = A.Keys
If Not AyIsEmpty(K) Then
    For Each I In K
        Push Dry, Array(I, A(I))
    Next
End If
Dim O As Drs
O.Fny = SplitSpc("Key Val")
O.Dry = Dry
DicDrs = O
End Function
