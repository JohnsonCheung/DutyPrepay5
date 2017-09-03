Attribute VB_Name = "Md_Srt__Tst"
Option Compare Database
Option Explicit

Private Sub Assert(BefSrt$(), AftSrt$())
If JnCrLf(BefSrt) = JnCrLf(AftSrt) Then Exit Sub
Dim A$(), I
A = AyMinus(BefSrt, AftSrt)
If Sz(A) = 0 Then Exit Sub
For Each I In A
    If I <> "" Then Stop
Next
End Sub

Sub MdSrt__Tst()
Dim I, Md As CodeModule
Dim BefSrt$(), AftSrt$()
For Each I In PjMdAy
    Set Md = I
'    If MdNm(Md) = "DaoDb" Then
        Debug.Print MdNm(Md)
        BefSrt = MdBdyLy(Md)
        AftSrt = SplitCrLf(MdSrtedBdyLines(Md))
        Assert BefSrt, AftSrt
'    End If
Next
End Sub
