Attribute VB_Name = "Ide_Win"
Option Compare Database
Option Explicit

Function CurCdWin() As VBIDE.Window
Set CurCdWin = Vbe.ActiveCodePane.Window
End Function

Function WinAy() As VBIDE.Window()
Dim O() As VBIDE.Window, W As VBIDE.Window
For Each W In Vbe.Windows
    PushObj O, W
Next
WinAy = O
End Function

Function WinAyOfCd() As VBIDE.Window()
WinAyOfCd = WinAyOfTy(vbext_wt_CodeWindow)
End Function

Function WinAyOfTy(T As vbext_WindowType) As VBIDE.Window()
WinAyOfTy = ObjAySelPrp(WinAy, "Type", T)
End Function

Sub WinClsCd(Optional ExceptMdNm$)
Dim I, W As VBIDE.Window
For Each I In WinAyOfCd
    Set W = I
    If WinMdNm(W) <> ExceptMdNm Then
        W.Close
    End If
Next
End Sub

Function WinMdNm$(A As VBIDE.Window)
WinMdNm = TakBet(A.Caption, " - ", " (Code)")
End Function
