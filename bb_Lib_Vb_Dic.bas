Attribute VB_Name = "bb_Lib_Vb_Dic"
Option Compare Database
Option Explicit
Function DicBrw(A As Dictionary)
DicDrs(A).Brw
End Function
Function DicDrs(A As Dictionary) As Drs
Dim Dry As New Dry, I
Dim K(): K = A.Keys
If Not AyIsEmpty(K) Then
    For Each I In K
        Dry.Push Array(I, A(I))
    Next
End If
Dim Fny$(): Fny = SplitSpc("Key Val")
DicDrs = Drs(Fny, Dry)
End Function
