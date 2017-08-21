VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Aql"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Private Sql$
Private Cn As ADODB.Connection
Property Get V() As AqlVal
V.Sql = Sql
Set V.Cn = Cn
End Property
Function Init(Sql_, Optional Cn_ As ADODB.Connection) As Aql
If IsEmpty(Sql) Then Stop
Sql = Sql_
Set Cn = DftCn(Cn_)
Set Init = Me
End Function

Private Sub Drs__Tst()
Const Fb$ = "N:\SapAccessReports\DutyPrepay5\DutyPrepay5.accdb"
Dim Sql$
Sql = "Select * from Permit"
'Dim Cn As ADODB.Connection: Set Cn = FxCnn("N:\SapAccessReports\DutyPrepay5\SAPDownloadExcel\KE24 2010-01c.xls")
Sql = "Select * from [Sheet1$]"
Drs(Fb, Sql).Brw
End Sub

Sub Run()
Cn.Execute Sql
End Sub
Function Brk(BrkColNm$, Optional MaxColWdt% = 100) As BrkLin
Set Brk = Drs.Brk(BrkColNm, MaxColWdt)
End Function
Function Drs() As Drs
Set Drs = ARs.Drs
End Function
Sub Brw(Optional MaxColWdt% = 100)
Drs.LyX(MaxColWdt).Brw
End Sub
Function Str$()
Str = Var
End Function
Function Int_%()
Int_ = Var
End Function

Function Lng&()
Lng = Var
End Function
Function ARs() As ARs
Set ARs = Nw.ARs(Rs)
End Function
Function Rs() As ADODB.Recordset
Dim O As New ADODB.Recordset
O.Open Sql, Cn
Set Rs = O
End Function

Function Var()
With Rs
    Var = .Fields(0).Value
    .Close
End With
End Function
