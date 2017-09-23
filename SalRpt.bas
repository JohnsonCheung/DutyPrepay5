Attribute VB_Name = "SalRpt"
Option Compare Database
Option Explicit
Public SR_PrmNm$
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

Function SR_CrdPfxTyDry() As Variant()
Static X As Boolean, Y
If Not X Then
    X = True
    Dim O()
    Push O, Array("134234", 1)
    Push O, Array("12323", 1)
    Push O, Array("2444", 2)
    Push O, Array("2443434", 2)
    Push O, Array("24424", 2)
    Push O, Array("3", 3)
    Push O, Array("5446561", 4)
    Push O, Array("6234341", 5)
    Push O, Array("6234342", 5)
    Y = O
End If
SR_CrdPfxTyDry = Y
End Function

Sub SR_Edt()
FtBrw PrmFt
End Sub

Function SR__MulSql$()
Dim O$()
Push O, ZMulSql_Drp
Push O, ZMulSql_T
Push O, ZMulSql_O
SR__MulSql = JnCrLf(O)
End Function

Function SR_Prm() As SR_Prm
If SR_PrmNm = "" Then
    SR_Prm = DftPrm
    Exit Function
End If
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

Private Function DftPrm() As SR_Prm
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
DftPrm = O
End Function

Private Function PrmDic() As Dictionary
Set PrmDic = DicByFt(PrmFt)
End Function

Private Function PrmFt$()
PrmFt = TstResPth & FmtQQ("SalRpt?.Sql3", SR_PrmNm)
End Function

Private Function ZMulSql_Drp$()
ZMulSql_Drp = MulSqlDrp("#Tx #TxMbr #MbrDta #Div #Sto #Crd #Cnt #Oup #MbrWs")
End Function

Private Function ZMulSql_O$()
Dim O$()
Push O, SR_O_Cnt_Sql
Push O, SR_O_Oup_Sql
Push O, SR_O_MbrWsOpt_Sql
O = AyRmvEmpty(O)
ZMulSql_O = JnDblCrLf(O)
End Function

Private Function ZMulSql_T$()
Dim O$()
Push O, SR_T_Tx_Sql
Push O, SR_T_UpdTx_Sql
Push O, SR_T_TxMbr_Sql
Push O, SR_T_MbrDtaOpt_Sql
Push O, SR_T_Div_Sql
Push O, SR_T_Sto_Sql
Push O, SR_T_Crd_Sql
ZMulSql_T = JnDblCrLf(O)
End Function

Private Sub PrmFt__Tst()
Debug.Print PrmFt
End Sub
