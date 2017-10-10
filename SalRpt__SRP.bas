Attribute VB_Name = "SalRpt__SRP"
Option Compare Database
Option Explicit
Public Const SRPNmLvs$ = _
"DivLis    " & _
"CrdLis    " & _
"StoLis    " & _
"BrkDiv    " & _
"BrkSto    " & _
"BrkCrd    " & _
"BrkMbr    " & _
"SumLvl    " & _
"InclNm    " & _
"InclAdr   " & _
"InclPhone " & _
"InclEmail " & _
"FmDte     " & _
"ToDte     "
Type SRP
'Prm
'    DivLis . 01 02 03
'    CrdLis . 1 2 3 4
'    StoLis . 001 002 003 004
'    ?BrkDiv . 1
'    ?BrkSto . 1
'    ?BrkCrd . 1
'    ?BrkMbr . 0
'    ?InclNm . 1O.Add "InclNm", True
'    ?InclAdr  . 1
'    ?InclPhone . 1
'    ?InclEmail . 1
'    SumLvl .Y
'    Fm . 20170101
'    To . 20170131
    DivLis As String
    CrdLis As String
    StoLis As String
    BrkDiv As Boolean
    BrkSto As Boolean
    BrkCrd As Boolean
    BrkMbr As Boolean
    SumLvl As String
    InclNm As Boolean
    InclAdr As Boolean
    InclPhone As Boolean
    InclEmail As Boolean
    FmDte As String
    ToDte As String
End Type

Function DftSRPD() As Dictionary
Dim X As Boolean, Y As New Dictionary
If Not X Then
    X = True
    With Y
        .Add "BrkCrd", False
        .Add "BrkDiv", False
        .Add "BrkMbr", False
        .Add "BrkSto", False
        .Add "CrdLis", ""
        .Add "StoLis", ""
        .Add "DivLis", ""
        .Add "FmDte", "20170101"
        .Add "ToDte", "20170131"
        .Add "SumLvl", "M"
        .Add "InclAdr", False
        .Add "InclNm", False
        .Add "InclPhone", False
        .Add "InclEmail", False
    End With
End If
Set DftSRPD = Y
End Function

Function SRP(Optional SRPNm$) As SRP
Dim D As Dictionary
    Set D = DicByLy(SRPLy(SRPNm))
SRPDAssertIsVdt D
Dim O As SRP
With O
    .BrkCrd = D("BrkCrd")
    .BrkDiv = D("BrkDiv")
    .BrkMbr = D("BrkMbr")
    .BrkSto = D("BrkSto")
    .CrdLis = D("CrdLis")
    .StoLis = D("StoLis")
    .DivLis = D("DivLis")
    .FmDte = D("FmDte")
    .ToDte = D("ToDte")
    .SumLvl = D("SumLvl")
    .InclNm = D("InclNm")
    .InclEmail = D("InclNm")
    .InclPhone = D("InclPhone")
    .InclAdr = D("InclAdr")
End With
SRP = O
End Function

Sub SRPDlt(Optional SRPNm$)
FfnDlt SRPFt(SRPNm)
End Sub

Sub SRPDmp(Optional SRPNm$)
Debug.Print "**PrmNm=" & SRPNm
Debug.Print "**PrmFt=" & SRPFt(SRPNm)
AyDmp SRPLy(SRPNm)
End Sub

Sub SRPEdt(SRPNm$)
FtBrw SRPFt(SRPNm)
End Sub

Sub SRPEns(Optional SRPNm$)
Dim Ft$
Ft = SRPFt(SRPNm)
If FfnIsExist(Ft) Then Exit Sub
AyWrt DicEqLy(DftSRPD), Ft
End Sub

Function SRPFt$(Optional SRPNm$)
SRPFt = SRPPth & FmtQQ("SalRpt-Prm?.txt", PrependDash(SRPNm))
End Function

Function SRPLy(Optional SRPNm$) As String()
SRPLy = FtLy(SRPFt(SRPNm))
End Function

Function SRPNy() As String()
SRPNy = PthFnAy(SRPPth, "*-Prm.txt")
End Function

Function SRPPth$()
SRPPth = TstResPth
End Function

Private Function DftSRP() As SRP
Dim O As SRP
With O
    .DivLis = "01 02 03"
    .CrdLis = "1 2 3 4"
    .StoLis = "001 002 003"
    .BrkDiv = True
    .BrkSto = True
    .BrkCrd = True
    .BrkMbr = True
    .SumLvl = "Y"
    .FmDte = "20170101"
    .ToDte = "20170131"
    .InclNm = True
    .InclAdr = True
    .InclPhone = True
    .InclEmail = True
End With
DftSRP = O
End Function
