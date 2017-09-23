Attribute VB_Name = "Ide_ChgTstMthMdy"
Option Compare Database
Option Explicit

Sub MdChgTstSubMdy(ToMdy$, Optional A As CodeModule)
Dim MthNm
For Each MthNm In MdTstSubNy(A)
    MthChgMdy MthNm, ToMdy, A
Next
End Sub

Sub MdChgTstSubToPrv(Optional A As CodeModule)
MdChgTstSubMdy "Private", A
End Sub

Sub MdChgTstSubToPub(Optional A As CodeModule)
MdChgTstSubMdy "Public", A
End Sub

Sub MthChgMdy(MthNm, ToMdy$, Optional A As CodeModule)
Dim Md As CodeModule: Set Md = DftMd(A)
Dim LnoAy%(): LnoAy = MdMthLnoAy(MthNm, A)
If AyIsEmpty(LnoAy) Then Stop
Dim Lno, NewLin$
For Each Lno In LnoAy
    NewLin = SrcLinRplMthMdy(Md.Lines(Lno, 1), ToMdy)
    MdRplLin CLng(Lno), NewLin, Md
Next
End Sub

Sub PjChgTstSubMdy(ToMdy$, Optional A As VBProject)
Dim I, Md As CodeModule
For Each I In PjMdAy(A)
    Set Md = I
    MdChgTstSubMdy ToMdy, Md
Next
End Sub

Sub PjChgTstSubToPrv(Optional A As VBProject)
Dim I, Md As CodeModule
For Each I In PjMdAy(A)
    Set Md = I
    MdChgTstSubToPrv Md
Next
End Sub
