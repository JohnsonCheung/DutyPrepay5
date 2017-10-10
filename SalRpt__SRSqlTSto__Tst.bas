Attribute VB_Name = "SalRpt__SRSqlTSto__Tst"
Option Compare Database
Option Explicit
Private Type TstDta
    StoLis As String
    BrkSto As Boolean
    Exp As String
End Type

Private Function TstDta1() As TstDta
With TstDta1
    .BrkSto = False
    .StoLis = "001 002"
    .Exp = ""
End With
End Function

Private Function TstDta2() As TstDta
With TstDta2
    .BrkSto = True
    .StoLis = "001 002"
    .Exp = "Select|    '0'+Loc_Code      Sto   ,|    Loc_Name          StoNm ,|    Loc_CName         StoCNm|  Into #Sto|  From Location|  Where '0'+Loc_Code in ('001','002')"
End With
End Function

Private Function TstDta3() As TstDta
With TstDta3
    .BrkSto = True
    .StoLis = ""
    .Exp = "Select|    '0'+Loc_Code      Sto   ,|    Loc_Name          StoNm ,|    Loc_CName         StoCNm|  Into #Sto|  From Location"
End With
End Function

Private Sub Tstr(A As TstDta)
With A
AssertActEqExp SR_SqlTSto(.BrkSto, .StoLis), .Exp
End With
End Sub

Private Sub SR_SqlTSto__Tst()
Tstr TstDta1
Tstr TstDta2
Tstr TstDta3
End Sub
