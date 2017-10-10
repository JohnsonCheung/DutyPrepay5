Attribute VB_Name = "SalRpt__SRSqlTTx__Tst"
Option Compare Database
Option Explicit
Private Type TstDta
    P As SRP
    CrdTyPfxDry() As Variant
    Exp As String
End Type

Private Function TstDta0() As TstDta

End Function

Private Function TstDta1() As TstDta
Dim O As TstDta
With O
    With .P
        .BrkCrd = True
        .BrkDiv = True
        .BrkMbr = True
        .BrkSto = True
        .CrdLis = "1 2 3"
        .DivLis = "01 02 03"
        .StoLis = "001 002 004"
        .FmDte = "20170101"
        .ToDte = "20170131"
        .SumLvl = "D"
    End With
    .Exp = _
        "Select|    Case When|      SHMCode Like '134234%' OR|      SHMCode Like '12323%'  THEN 1|      Else Case When|      SHMCode Like '2444%'    OR|      SHMCode Like '2443434%' OR|      SHMCode Like '24424%'   THEN 2|      Else Case When|      SHMCode Like '3%' THEN 3|      Else 4|      End End End                                                              Crd  ,|    Sum(SHAmount)                                                            Amt  ,|    Sum(SHQty)                                                               Qty  ,|    Count(SHInvoice + SHSDate + SHRef)                                       Cnt  ,|    Mbr-Expr                                                                 Mbr  ,|    Div-Expr                                                                 Div  ,|    Sto-Expr                                                                 Sto  ,|    SUBSTR(SHSDate,1,4)                                                      TxY  ,|    SUBSTR(SHSDate,5,2)" & _
        "TxM  ,|    TxW-Expr                                                                 TxW  ,|    SUBSTR(SHSDate,7,2)                                                      TxD  ,|    SUBSTR(SHSDate,1,4)+'/'+SUBSTR(SHSDate,5,2)+'/'+SUBSTR(SHSDate,7,2)      TxDte|  Into #Tx|  From SaleHistory|  Where SHDate Between '20170101' and '20170131'|    And Case When|SHMCode Like '134234%' OR|SHMCode Like '12323%'  THEN 1|Else Case When|SHMCode Like '2444%'    OR|SHMCode Like '2443434%' OR|SHMCode Like '24424%'   THEN 2|Else Case When|SHMCode Like '3%' THEN 3|Else 4|End End End  in (1,2,3)|    And Div-Expr in ('01','02','03')|    And Sto-Expr in ('001','002','004')|  Group By|Case When|SHMCode Like '134234%' OR|SHMCode Like '12323%'  THEN 1|Else Case When|SHMCode Like '2444%'    OR|SHMCode Like '2443434%' OR|SHMCode Like '24424%'   THEN 2|Else Case When|SHMCode Like '3%' THEN 3|Else 4|End End End                                                         ,|Mbr-Expr" & _
        ",|Div-Expr                                                            ,|Sto-Expr                                                            ,|SUBSTR(SHSDate,1,4)                                                 ,|SUBSTR(SHSDate,5,2)                                                 ,|SUBSTR(SHSDate,7,2)                                                 ,|SUBSTR(SHSDate,1,4)+'/'+SUBSTR(SHSDate,5,2)+'/'+SUBSTR(SHSDate,7,2)"
End With
TstDta1 = O
End Function

Private Function TstDta2() As TstDta
Dim O As TstDta
With O
    With .P
        .BrkCrd = True
        .BrkDiv = True
        .BrkMbr = True
        .BrkSto = True
        .CrdLis = "1 2 3"
        .DivLis = "01 02 03"
        .StoLis = "001 002 004"
        .FmDte = "20170101"
        .ToDte = "20170131"
        .SumLvl = "D"
    End With
    .Exp = ""
End With
TstDta2 = O
End Function

Private Function TstDta3() As TstDta
With TstDta3
    With .P
    End With
    .Exp = ""
End With
End Function

Private Function TstDtaAy() As TstDta()
Dim O() As TstDta
TstDtaPush O, TstDta0
End Function

Private Sub TstDtaPush(O() As TstDta, I As TstDta)

End Sub

Private Sub Tstr(A As TstDta)
Dim ECrd$
With A
    ECrd = SRECrd(.P.CrdLis, .CrdTyPfxDry)
    AssertActEqExp SR_SqlTTx(A.P, ECrd), .Exp
End With
End Sub

Private Sub SR_SqlTSto__Tst()
Dim Ay() As TstDta
Dim J%
For J = 0 To UBound(Ay)
    Tstr Ay(J)
Next
End Sub

