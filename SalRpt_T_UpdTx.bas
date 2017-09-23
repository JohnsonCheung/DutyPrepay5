Attribute VB_Name = "SalRpt_T_UpdTx"
Option Compare Database
Option Explicit
Const ETxWD$ = ""
Function SR_T_UpdTx_Sql$()
SR_T_UpdTx_Sql = FmtQQ("Update #Tx|  Set|  TxWD =|?", ETxWD)
End Function


