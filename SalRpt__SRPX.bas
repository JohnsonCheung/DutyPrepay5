Attribute VB_Name = "SalRpt__SRPX"
Option Compare Database
Option Explicit
Type SRPX
    ECrd As String
    InDiv As String
    InSto As String
    InCrd As String
    InclFldTxD As Boolean
    InclFldTxY As Boolean
    InclFldTxM As Boolean
    InclFldTxW As Boolean
End Type

Function SRPX(P As SRP, ECrd$) As SRPX
Dim InclFldTxD As Boolean
Dim InclFldTxM As Boolean
Dim InclFldTxW As Boolean
Dim InclFldTxY As Boolean
    Dim SumLvl$
    SumLvl = P.SumLvl
    Select Case SumLvl
    Case "D": InclFldTxD = True
    End Select
    '
    Select Case SumLvl
    Case "D", "W", "M": InclFldTxM = True
    End Select
    '
    Select Case SumLvl
    Case "D", "W": InclFldTxW = True
    End Select
    '
    Select Case SumLvl
    Case "D", "W", "M", "Y": InclFldTxY = True
    End Select

Dim O As SRPX
With O
    .ECrd = ECrd
    .InclFldTxD = InclFldTxD
    .InclFldTxM = InclFldTxM
    .InclFldTxW = InclFldTxW
    .InclFldTxY = InclFldTxY
    .InCrd = JnComma(SplitLvs(P.CrdLis))
    .InDiv = JnComma(AyQuoteSng(SplitLvs(P.DivLis)))
    .InSto = JnComma(AyQuoteSng(SplitLvs(P.StoLis)))
End With
SRPX = O
End Function

Function SRPXDic(A As SRPX) As Dictionary

End Function
