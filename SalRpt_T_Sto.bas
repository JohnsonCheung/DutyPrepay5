Attribute VB_Name = "SalRpt_T_Sto"
Option Compare Database
Option Explicit
Private Prm As SR_Prm
Function SR_T_Sto_Sql$()
Prm = SR_Prm
SR_T_Sto_Sql = FmtQQ("Select|?|  Into #Sto|  From Loc", Fld)
End Function
Private Function Fld$()

End Function

