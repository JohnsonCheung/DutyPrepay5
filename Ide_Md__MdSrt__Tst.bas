Attribute VB_Name = "Ide_Md__MdSrt__Tst"
Option Explicit
Option Compare Database

Private Sub Assert(BefSrt$(), AftSrt$())
If JnCrLf(BefSrt) = JnCrLf(AftSrt) Then Exit Sub
Dim A$(), I
A = AyMinus(BefSrt, AftSrt)
If Sz(A) = 0 Then Exit Sub
For Each I In A
    If I <> "" Then Stop
Next
Stop
End Sub

Sub MdSrt__Tst()
Dim I, Md As CodeModule
Dim BefSrt$(), AftSrt$()
For Each I In PjMdAy
    Set Md = I
    If MdNm(Md) = "Vb" Then
        Debug.Print MdNm(Md)
        BefSrt = MdBdyLy(Md)
        AftSrt = SplitCrLf(MdSrtedBdyLines(Md))
        AyBrw AftSrt
        If Not AyIsEmpty(AftSrt) Then
            If AyLasEle(AftSrt) = "" Then
                AyBrw AftSrt
                Stop
            End If
        End If
        Assert BefSrt, AftSrt
    End If
Next
End Sub

Sub Tst()
MdSrt__Tst
End Sub
