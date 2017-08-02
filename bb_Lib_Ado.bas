Attribute VB_Name = "bb_Lib_Ado"
Option Compare Database
Option Explicit
Function ASqlLng&(Cn As ADODB.Connection, Sql$)
ASqlLng = ASqlV(Cn, Sql)
End Function
Function ASqlV(Cn As ADODB.Connection, Sql$)
Dim Rs As New ADODB.Recordset
With ASqlRs(Cn, Sql)
    ASqlV = .Fields(0).Value
    .Close
End With
End Function
Function ASqlRs(Cn As ADODB.Connection, Sql) As ADODB.Recordset
Dim O As New ADODB.Recordset
O.Open Sql, Cn
Set ASqlRs = O
End Function
Sub ASqlDrs__Tst()
Dim Cn As ADODB.Connection: Set Cn = FxAdoCnn("N:\SapAccessReports\DutyPrepay5\SAPDownloadExcel\KE24 2010-01c.xls")
Dim Sql$: Sql = "Select * from [Sheet1$]"
Dim Drs As Drs: Drs = ASqlDrs(Cn, Sql)
BrwDrs Drs
End Sub
Function ASqlDrs(Cn As ADODB.Connection, Sql$) As Drs
ASqlDrs = ARsDrs(ASqlRs(Cn, Sql))
End Function
Sub RunASql(Cn As ADODB.Connection, Sql)
Cn.Execute Sql
End Sub
Sub RunASqlAy(Cn As ADODB.Connection, SqlAy$())
If IsEmptyAy(SqlAy) Then Exit Sub
Dim Sql
For Each Sql In SqlAy
    RunASql Cn, Sql
Next
End Sub
Sub FxAdoCnn__Tst()
Dim A As ADODB.Connection
Set A = FxAdoCnn("N:\SapAccessReports\DutyPrepay5\SAPDownloadExcel\KE24 2010-01c.xls")
Stop
End Sub
Function FxAdoCnn(Fx) As ADODB.Connection
Dim O As New ADODB.Connection
'Provider=Microsoft.ACE.OLEDB.12.0;Data Source=c:\myFolder\myOldExcelFile.xls;
'Extended Properties="Excel 8.0;HDR=YES";
O.ConnectionString = FmtQQ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=?;Extended Properties=""Excel 8.0;HDR=YES""", Fx)
O.Open
Set FxAdoCnn = O
End Function
Sub FxAdoCat__Tst()
Dim A As ADOX.Catalog
Set A = FxAdoCat("N:\SapAccessReports\DutyPrepay5\SAPDownloadExcel\KE24 2010-01c.xls")
Stop
End Sub
Sub FxWsNy__Tst()
DmpAy FxWsNy("N:\SapAccessReports\DutyPrepay5\SAPDownloadExcel\KE24 2010-01c.xls")
End Sub
Function FxWsNy(Fx) As String()
Dim T As ADOX.Table
Dim O$()
For Each T In FxAdoCat(Fx).Tables
    Push O, RmvLasNChr(T.Name)
Next
FxWsNy = O
End Function
Sub FxWsFny__Tst()
DmpAy FxWsFny("N:\SapAccessReports\DutyPrepay5\SAPDownloadExcel\KE24 2010-01c.xls", "Sheet1")
End Sub
Function FxWsFny(Fx, WsNm$) As String()
Dim Cat As ADOX.Catalog: Set Cat = FxAdoCat(Fx)
Dim C As ADOX.Column
Dim O$()
Dim N$: N = WsNm & "$"
For Each C In Cat.Tables(N).Columns
    Push O, C.Name
Next
FxWsFny = O
End Function
Function FxAdoCat(Fx) As Catalog
Dim O As New Catalog
Set O.ActiveConnection = FxAdoCnn(Fx)
Set FxAdoCat = O
End Function
