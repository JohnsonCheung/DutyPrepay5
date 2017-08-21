VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ado_Cn_CnStr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Function ByFb$(Fb)
ByFb = FmtQQ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=?;", Fb)
End Function
Function ByFx$(Fx)
ByFx = FmtQQ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=?;Extended Properties=""Excel 8.0;HDR=YES""", Fx)
End Function

