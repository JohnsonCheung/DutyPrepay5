Attribute VB_Name = "SalRpt"
Option Compare Database
Option Explicit
'SRE_ = SalRpt-Expression
Public Const SRE_TxWD$ = _
"CASE WHEN TxWD1 = 1 then 'Sun'" & _
"|ELSE WHEN TxWD1 = 2 THEN 'Mon'" & _
"|ELSE WHEN TxWD1 = 3 THEN 'Tue'" & _
"|ELSE WHEN TxWD1 = 4 THEN 'Mon'" & _
"|ELSE WHEN TxWD1 = 5 THEN 'Thu'" & _
"|ELSE WHEN TxWD1 = 6 THEN 'Fri'" & _
"|ELSE WHEN TxWD1 = 7 THEN 'Sat'" & _
"|ELSE Null" & _
"|END END END END END END END"

Sub SalRpt_Stop()
Stop
End Sub

Function SR__MulSql$()
Dim O$()
Push O, ZMulSql_Drp
Push O, ZMulSql_T
Push O, ZMulSql_O
SR__MulSql = JnCrLf(O)
End Function

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

Sub SR_DltPrmFt()

End Sub

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
