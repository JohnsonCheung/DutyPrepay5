VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ado_Cn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private CnStr As Ado_Cn_CnStr
Function ByFb(Fb) As ACn
Set ByFb = NewCn(CnStr.ByFb(Fb))
End Function
Function ByFx(Fx) As ADODB.Connection
Set ByFx = NewCn(CnStr.ByFx(Fx))
End Function
Private Function NewCn(CnStr$) As ADODB.Connection
Dim O As New ADODB.Connection
O.ConnectionString = CnStr
O.Open
Set NewCn = O
End Function
