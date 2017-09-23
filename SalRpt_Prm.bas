Attribute VB_Name = "SalRpt_Prm"
Option Compare Database
Option Explicit
Private SR_PrmNm_$
Type SR_Prm
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

Function SR_Prm() As SR_Prm
Dim D As Dictionary
    Set D = PrmDic
Dim O As SR_Prm
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
SR_Prm = O
End Function

Sub SR_Prm_Stop()
Stop
End Sub

Sub SR_PrmDlt()
FfnDlt PrmFt
End Sub

Sub SR_PrmDmp()
Debug.Print "**PrmNm=" & SR_PrmNm
Debug.Print "**PrmFt=" & PrmFt
DicDmp PrmDic
End Sub

Sub SR_PrmEdt()
FtBrw PrmFt
End Sub

Property Get SR_PrmNm$()
SR_PrmNm = SR_PrmNm_
End Property

Property Let SR_PrmNm(V$)
SR_PrmNm_ = V
Dim D As Dictionary
Set D = PrmDic(V)
With SR_Prm
    
End With
End Property

Function SRS_InclFldDiv() As Boolean
SRS_InclFldDiv = SR_Prm.BrkDiv
End Function

Function SRS_InclFldMbr() As Boolean
SRS_InclFldMbr = SR_Prm.BrkMbr
End Function

Function SRS_InclFldSto() As Boolean
SRS_InclFldSto = SR_Prm.BrkDiv
End Function

Function SRS_InclFldTxD() As Boolean
Select Case SR_Prm.SumLvl
Case "D": SRS_InclFldTxD = True
End Select
End Function

Function SRS_InclFldTxDte() As Boolean
SRS_InclFldTxDte = SRS_InclFldTxD
End Function

Function SRS_InclFldTxM() As Boolean
Select Case SR_Prm.SumLvl
Case "D", "W", "M": SRS_InclFldTxM = True
End Select
End Function

Function SRS_InclFldTxW() As Boolean
Select Case SR_Prm.SumLvl
Case "D", "W": SRS_InclFldTxW = True
End Select
End Function

Function SRS_InclFldTxY() As Boolean
Select Case SR_Prm.SumLvl
Case "D", "W", "M", "Y": SRS_InclFldTxY = True
End Select
End Function

Function SRS_InclSelCrd() As Boolean
SRS_InclSelCrd = SR_Prm.CrdLis <> ""
End Function

Function SRS_InclSelDiv() As Boolean
SRS_InclSelDiv = SR_Prm.DivLis <> ""
End Function

Function SRS_InclSelSto() As Boolean
SRS_InclSelSto = SR_Prm.StoLis <> ""
End Function

Private Sub AssertIsPrmDic(A As Dictionary)
DicAssertKeyLvs A, "DivLis CrdLis StoLis Brkdiv BrkSto BrkCrd BrkMbr SumLvl FmDte TDte InclNm InclAdr InclPhone InclEmail"
End Sub

Private Function DftPrmDic() As Dictionary
Dim X As Boolean, Y As New Dictionary
If Not X Then
    X = True
    With Y
        .Add "BrkCrd", False
        .Add "BrkDiv", False
        .Add "BrkMbr", False
        .Add "BrkSto", ""
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
Set DftPrmDic = Y
End Function

Private Function InitPrm() As SR_Prm
Dim O As SR_Prm
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
InitPrm = O
End Function

Private Function PrmDic() As Dictionary
Dim O As Dictionary
Dim Ft$: Ft = PrmFt
Set O = DicByFt(Ft)
AssertIsPrmDic O
Set PrmDic = O
End Function

Private Function PrmFt$()
PrmFt = XPrmFt(SR_PrmNm)
End Function

Private Function XPrmFt$(PrmNm$)
Dim O$: O = TstResPth & FmtQQ("SalRpt-Prm?.txt", PrependDash(PrmNm))
If Not Fso.FileExists(O) Then
    AyWrt DicKVLy(DftPrmDic), O
End If
XPrmFt = O
End Function

Private Sub PrmFt__Tst()
Debug.Print PrmFt
End Sub
