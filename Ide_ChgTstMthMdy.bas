Attribute VB_Name = "Ide_ChgTstMthMdy"
Option Compare Database
Option Explicit

Sub MdChgTstMthMdy(ToMdy$, Optional A As CodeModule)
Dim MthNm
For Each MthNm In MdTstMthNy(A)
    MthChgMdy MthNm, ToMdy, A
Next
End Sub

Sub MdChgTstMthToPrv(Optional A As CodeModule)
MdChgTstMthMdy "Private", A
End Sub

Sub MdChgTstMthToPub(Optional A As CodeModule)
MdChgTstMthMdy "Public", A
End Sub

Sub MthChgMdy(MthNm, ToMdy$, Optional A As CodeModule)
Dim Md As CodeModule: Set Md = DftMd(A)
Dim LnoAy%(): LnoAy = MdMthLnoAy(MthNm, A)
If AyIsEmpty(LnoAy) Then Stop
Dim Lno, NewLin$
For Each Lno In LnoAy
    NewLin = MthLinRplMdy(Md.Lines(Lno, 1), ToMdy)
    MdRplLin CLng(Lno), NewLin, Md
Next
End Sub

Sub PjChgTstMthMdy(ToMdy$, Optional A As Vbproject)
Dim I, Md As CodeModule
For Each I In PjMdAy(, A)
    Set Md = I
    MdChgTstMthMdy ToMdy, Md
Next
End Sub

Sub PjChgTstMthToPrv(Optional A As Vbproject)
Dim I, Md As CodeModule
For Each I In PjMdAy(, A)
    Set Md = I
    MdChgTstMthToPrv Md
Next
End Sub
