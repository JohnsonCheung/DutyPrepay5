Attribute VB_Name = "SalRpt_T_Crd"
Option Compare Database
Option Explicit
Private Prm As SR_Prm

Function SR_T_Crd_Sql$()
Prm = SR_Prm
SR_T_Crd_Sql = FmtQQ("Select|?|  Into #Crd|  From Division", Fld)
End Function

Private Function Fld$()

End Function
