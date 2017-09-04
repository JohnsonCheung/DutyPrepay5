Attribute VB_Name = "Vb_Oy"
Option Explicit
Option Compare Database

Function OyPrp(Oy, PrpNm$, Optional Oup)
Dim O
    If Not IsMissing(Oup) Then
        O = Oup
        Erase O
    Else
        O = EmptyAy
    End If
    If AyIsEmpty(Oy) Then GoTo X
    Dim I
    For Each I In Oy
        Push O, CallByName(I, PrpNm, VbGet)
    Next
X:
    OyPrp = O
End Function

Private Sub OyPrp__Tst()
Dim CdPanAy() As CodePane
CdPanAy = OyPrp(PjMdAy, "CodePane", CdPanAy)
Stop
End Sub

Sub Tst()
OyPrp__Tst
End Sub
