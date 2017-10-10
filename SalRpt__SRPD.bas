Attribute VB_Name = "SalRpt__SRPD"
Option Compare Database
Option Explicit

Sub SRPDAssertIsVdt(A As Dictionary)
'SRPD = Sales Report Parameter Dictionary
If Not SRPDIsVdt(A) Then Stop
End Sub

Function SRPDByNm(PrmNm$) As Dictionary
Dim O As Dictionary
Set O = DicByFt(SRPFt(PrmNm))
SRPDAssertIsVdt O
Set SRPDByNm = O
End Function

Function SRPDic(P As SRP) As Dictionary
Dim O As Dictionary
With O
    .Add "", P.BrkCrd
    .Add "", P.BrkDiv
    .Add "", P.BrkMbr
    .Add "", P.BrkSto
    .Add "", P.BrkCrd
    .Add "", P.InclNm
    .Add "", P.InclAdr
    .Add "", P.InclEmail
    .Add "", P.InclPhone
    .Add "", P.CrdLis
    .Add "", P.DivLis
    .Add "", P.StoLis
    .Add "", P.FmDte
    .Add "", P.ToDte
    .Add "", P.SumLvl
End With
If O.Count <> Sz(SplitLvs(SRPNmLvs)) Then Stop
Set SRPDic = O
End Function

Function SRPDIsVdt(A As Dictionary) As Boolean
SRPDIsVdt = DicHasKeyLvs(A, SRPNmLvs)
End Function
