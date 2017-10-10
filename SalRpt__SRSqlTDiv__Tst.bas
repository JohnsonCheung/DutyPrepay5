Attribute VB_Name = "SalRpt__SRSqlTDiv__Tst"
Option Compare Database
Option Explicit
Private Type TstDta
    BrkDiv As Boolean
    DivLis As String
    Exp As String
End Type

Private Function TstDta1() As TstDta
With TstDta1
    .BrkDiv = False
    .DivLis = "01 02"
End With
End Function

Private Function TstDta2() As TstDta
With TstDta2
    .BrkDiv = True
    .DivLis = "01 02"
    .Exp = "Select|    Dept + Division      Div   ,|    DivNm                DivNm ,|    Seq                  DivSeq,|    Status               DivSts|  Into #Div|  From Division|  Where Dept + Division in ('01','02')"
End With
End Function

Private Function TstDta3() As TstDta
With TstDta3
    .BrkDiv = True
    .DivLis = ""
    .Exp = "Select|    Dept + Division      Div   ,|    DivNm                DivNm ,|    Seq                  DivSeq,|    Status               DivSts|  Into #Div|  From Division"
End With
End Function

Private Sub Tstr(A As TstDta)
With A
    AssertActEqExp SR_SqlTDiv(.BrkDiv, .DivLis), .Exp
End With
End Sub

Private Sub SR_SqlTDiv__Tst()
Tstr TstDta1
Tstr TstDta2
Tstr TstDta3
End Sub
