Attribute VB_Name = "SalRpt_T_UpdTx"
Option Compare Database
Option Explicit

Function SR_T_UpdTx_Sql$()
SR_T_UpdTx_Sql = _
SqpUpd("#Tx") & _
SqpSet("TxWD", Array(SRE_TxWD))
End Function

Private Function ExprDic() As Dictionary
Dim O As New Dictionary
O.Add "TxWD", SRE_TxWD
Set ExprDic = O
End Function

