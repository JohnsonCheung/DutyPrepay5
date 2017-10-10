Attribute VB_Name = "SalRpt__SRSqlTCrd__Tst"
Option Compare Database
Option Explicit
Private Type TstDta
    CrdLis As String
    BrkCrd As Boolean
    Exp As String
End Type

Private Function TstDta1() As TstDta
With TstDta1
    .BrkCrd = False
    .CrdLis = "1 2"
    .Exp = ""
End With
End Function

Private Function TstDta2() As TstDta
With TstDta2
    .BrkCrd = True
    .CrdLis = "1 2"
    .Exp = "Select|    CrdTyId      Crd  ,|    CrdTyNm      CrdNm|  Into #Crd|  From JR_FrqMbrLis_#CrdTy()|  Where CrdTyId in (1,2)"
End With
End Function

Private Function TstDta3() As TstDta
With TstDta3
    .BrkCrd = True
    .CrdLis = ""
    .Exp = "Select|    CrdTyId      Crd  ,|    CrdTyNm      CrdNm|  Into #Crd|  From JR_FrqMbrLis_#CrdTy()"
End With
End Function

Private Sub Tstr(A As TstDta)
With A
    Dim Act$
    Act = SR_SqlTCrd(.BrkCrd, .CrdLis)
    AssertActEqExp Act, .Exp
End With
End Sub

Private Sub SR_SqlTCrd__Tst()
Tstr TstDta1
Tstr TstDta2
Tstr TstDta3
End Sub
