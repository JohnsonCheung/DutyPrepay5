VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ACn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Private Cn As ADODB.Connection
Property Get Val() As ADODB.Connection
Set Val = Cn
End Property

Function Drs(Sql) As Drs
Set Drs = Aql(Sql).Drs
End Function
Function Aql(Sql) As Aql
Set Aql = Nw.Aql(Sql, Cn)
End Function
Private Sub FxCat__Tst()
Dim A As ADOX.Catalog
'Set A = FxCat("N:\SapAccessReports\DutyPrepay5\SAPDownloadExcel\KE24 2010-01c.xls")
Stop
End Sub

Function WsFny(WsNm$) As String()
WsFny = ACat.Fny(WsNm & "$")
End Function
Private Sub WsFny__Tst()
'AyDmp FxWsFny("N:\SapAccessReports\DutyPrepay5\SAPDownloadExcel\KE24 2010-01c.xls", "Sheet1")
End Sub
Function Cat() As ACat
Dim O As New Catalog
Set O.ActiveConnection = Cn
Set Cat = ACat(O)
End Function
Function ACat() As ACat
Set ACat = Nw.ACat(Cat)
End Function
Function WsNy() As String()
WsNy = ACat.TnyX.RmvLasChr
End Function
Private Sub WsNy__Tst()
'AyDmp FxWsNy("N:\SapAccessReports\DutyPrepay5\SAPDownloadExcel\KE24 2010-01c.xls")
End Sub
Private Sub Class_Terminate()
On Error Resume Next
Cn.Close
End Sub
