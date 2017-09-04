Attribute VB_Name = "Ado"
Option Explicit
Option Compare Database

Function AFldsDr(AFlds As ADODB.Fields) As Variant()
Dim O()
ReDim O(AFlds.Count - 1)
Dim J%, F As ADODB.Field
For Each F In AFlds
    O(J) = F.Value
    J = J + 1
Next
AFldsDr = O
End Function

Function AFldsFny(AFlds As ADODB.Fields) As String()
Dim O$()
Dim F As ADODB.Field
For Each F In AFlds
    Push O, F.Name
Next
AFldsFny = O
End Function

Function AqlDrs(Cn As ADODB.Connection, Sql) As Drs
Dim Rs As ADODB.Recordset
Set Rs = AqlRs(Cn, Sql)
AqlDrs = ARsDrs(Rs)
Rs.Close
End Function

Function AqlLng&(Cn As ADODB.Connection, Sql$)
AqlLng = AqlV(Cn, Sql)
End Function

Function AqlRs(Cn As ADODB.Connection, Sql) As ADODB.Recordset
Dim O As New ADODB.Recordset
O.Open Sql, Cn
Set AqlRs = O
End Function

Function AqlV(Cn As ADODB.Connection, Sql$)
Dim Rs As New ADODB.Recordset
With AqlRs(Cn, Sql)
    AqlV = .Fields(0).Value
    .Close
End With
End Function

Function ARsDrs(Rs As ADODB.Recordset) As Drs
Dim O As Drs
O.Fny = ARsFny(Rs)
O.Dry = ARsDry(Rs)
ARsDrs = O
End Function

Function ARsDry(Rs As ADODB.Recordset) As Variant()
Dim O()
With Rs
    While Not .EOF
        Push O, AFldsDr(Rs.Fields)
        .MoveNext
    Wend
End With
ARsDry = O
End Function

Function ARsFny(Rs As ADODB.Recordset) As String()
ARsFny = AFldsFny(Rs.Fields)
End Function

Function CatCnn(A As Catalog) As ADODB.Connection
Set CatCnn = A.ActiveConnection
End Function

Function FbAqlDrs(Fb, Sql) As Drs
FbAqlDrs = AqlDrs(FbCnn(Fb), Sql)
End Function

Function FbCnn(Fb) As ADODB.Connection
Dim O As New ADODB.Connection
O.ConnectionString = FmtQQ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=?;", Fb)
O.Open
Set FbCnn = O
End Function

Function FxAqlDrs(Fx, Sql) As Drs
FxAqlDrs = AqlDrs(FxCnn(Fx), Sql)
End Function

Function FxCat(Fx) As Catalog
Dim O As New Catalog
Set O.ActiveConnection = FxCnn(Fx)
Set FxCat = O
End Function

Function FxCnn(Fx) As ADODB.Connection
Dim O As New ADODB.Connection
'Provider=Microsoft.ACE.OLEDB.12.0;Data Source=c:\myFolder\myOldExcelFile.xls;
'Extended Properties="Excel 8.0;HDR=YES";
O.ConnectionString = FmtQQ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=?;Extended Properties=""Excel 8.0;HDR=YES""", Fx)
O.Open
Set FxCnn = O
End Function

Function FxWsFny(Fx, WsNm$) As String()
Dim Cat As ADOX.Catalog: Set Cat = FxCat(Fx)
Dim C As ADOX.Column
Dim O$()
Dim N$: N = WsNm & "$"
For Each C In Cat.Tables(N).Columns
    Push O, C.Name
Next
FxWsFny = O
End Function

Function FxWsNy(Fx) As String()
Dim T As ADOX.Table
Dim O$()
For Each T In FxCat(Fx).Tables
    Push O, RmvLasNChr(T.Name)
Next
FxWsNy = O
End Function

Sub RunAql(Cn As ADODB.Connection, Sql)
Cn.Execute Sql
End Sub

Sub RunAqlAy(Cn As ADODB.Connection, SqlAy$())
If AyIsEmpty(SqlAy) Then Exit Sub
Dim Sql
For Each Sql In SqlAy
    RunAql Cn, Sql
Next
End Sub

Sub RunFbAql(Fb, Sql)
RunAql FbCnn(Fb), Sql
End Sub

Sub RunFxAql(Fx, Sql)
RunAql FxCnn(Fx), Sql
End Sub

Private Sub AqlDrs__Tst()
Dim Cn As ADODB.Connection: Set Cn = FxCnn("N:\SapAccessReports\DutyPrepay5\SAPDownloadExcel\KE24 2010-01c.xls")
Dim Sql$: Sql = "Select * from [Sheet1$]"
Dim Drs As Drs: Drs = AqlDrs(Cn, Sql)
DrsBrw Drs
End Sub

Private Sub FbAqlDrs__Tst()
Const Fb$ = "N:\SapAccessReports\DutyPrepay5\DutyPrepay5.accdb"
Const Sql$ = "Select * from Permit"
DrsBrw FbAqlDrs(Fb, Sql)
End Sub

Private Sub FbCnn__Tst()
Dim A As ADODB.Connection
Set A = FbCnn("N:\SapAccessReports\DutyPrepay5\DutyPrepay5_data.mdb")
Stop
End Sub

Private Sub FxAqlDrs__Tst()
Const Fx$ = "N:\SapAccessReports\DutyPrepay5\SAPDownloadExcel\KE24 2010-01c.xls"
Const Sql$ = "Select * from [Sheet1$]"
DrsBrw FxAqlDrs(Fx, Sql)
End Sub

Private Sub FxCat__Tst()
Dim A As ADOX.Catalog
Set A = FxCat("N:\SapAccessReports\DutyPrepay5\SAPDownloadExcel\KE24 2010-01c.xls")
Stop
End Sub

Private Sub FxCnn__Tst()
Dim A As ADODB.Connection
Set A = FxCnn("N:\SapAccessReports\DutyPrepay5\SAPDownloadExcel\KE24 2010-01c.xls")
Stop
End Sub

Private Sub FxWsFny__Tst()
AyDmp FxWsFny("N:\SapAccessReports\DutyPrepay5\SAPDownloadExcel\KE24 2010-01c.xls", "Sheet1")
End Sub

Private Sub FxWsNy__Tst()
AyDmp FxWsNy("N:\SapAccessReports\DutyPrepay5\SAPDownloadExcel\KE24 2010-01c.xls")
End Sub

Private Sub RunFbAql__Tst()
Const Fb$ = "N:\SapAccessReports\DutyPrepay5\tmp.accdb"
Const Sql$ = "Select * into [#a] from Permit"
RunFbAql Fb, Sql
End Sub

Private Sub RunFxAql__Tst()
Const Fx$ = "N:\SapAccessReports\DutyPrepay5\SAPDownloadExcel\KE24 2010-01c.xls"
Const Sql$ = "Select * into [Sheet21] from [Sheet1$]"
RunFxAql Fx, Sql
End Sub

Sub Tst()
AqlDrs__Tst
FbAqlDrs__Tst
FbCnn__Tst
FxAqlDrs__Tst
FxCat__Tst
FxCnn__Tst
FxWsFny__Tst
FxWsNy__Tst
RunFbAql__Tst
RunFxAql__Tst
End Sub
