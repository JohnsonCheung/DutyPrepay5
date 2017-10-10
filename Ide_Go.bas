Attribute VB_Name = "Ide_Go"
Option Compare Database

Sub MthGo(MthNm, Optional ClsOthMd As Boolean, Optional A As Vbproject)
WinClsCd
Dim I
For Each I In AySelPfx(PjMthDotNy(MthNm, A), "." & MthNm)
    'MthDotNmGo I, A
Next
End Sub
