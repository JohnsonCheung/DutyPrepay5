VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Fx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Public Fx$
Function HasWs(WsNm) As Boolean
HasWs = WsNyX.Has(WsNm)
End Function
Function WsNyX() As Ay
Set WsNyX = Ay(WsNy)
End Function
Function WsNy() As String()
Dim T As ADOX.Table
Dim O$()
For Each T In Cat.Tables
    Push O, RmvLasNChr(T.Name)
Next
WsNy = O
End Function

Function Cat() As Catalog
Dim O As New Catalog
Set O.ActiveConnection = Cnn
Set Cat = O
End Function

Function Cnn() As ADODB.Connection
Dim O As New ADODB.Connection
'Provider=Microsoft.ACE.OLEDB.12.0;Data Source=c:\myFolder\myOldExcelFile.xls;
'Extended Properties="Excel 8.0;HDR=YES";
O.ConnectionString = FmtQQ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=?;Extended Properties=""Excel 8.0;HDR=YES""", Fx)
O.Open
Set Cnn = O
End Function

Sub RmvWsIfExist(WsNm)
If Not HasWs(WsNm) Then Exit Sub
WbX.DltWs(WsNm).Sav.ClsNoSav
End Sub
Function WbX() As Wb
Set WbX = Nw.Wb(Wb)
End Function
Function Wb() As Workbook
Set Wb = Xls.Workbooks.Open(Fx)
End Function
Function Db(Optional WsNmMapStr$) As Db
Dim O As Db
    Set O = Tmp.Db
Dim mWsNy$()
Dim Tny$()
    If WsNmMapStr = "" Then
        mWsNy = WsNy
    Else
        With BrkMapStr(WsNmMapStr)
            mWsNy = .Sy2
            Tny = .Sy1
        End With
    End If
Db.LnkFx Fx, mWsNy, Tny
Set Db = O
End Function
Function Init(Fx_) As Fx
Fx = Fx_
Set Init = Me
End Function
Private Function ZZ() As Fx

End Function
Sub RmvWsIfExist__Tst()
Dim T$: T = Tmp.Fx
Dim Wb As Workbook
Set Wb = NewWb
Wb.Sheets.Add
Wb.SaveAs T
Dim WsNyBef$(), WsNyAft$()
    WsNyBef = ZZ.WsNy
    ZZ.RmvWsIfExist "Sheet1"
    WsNyAft = ZZ.WsNy
Dim Exp$()
    Exp = Ay(WsNyBef).Minus(Array("Sheet1"))
Ay(Exp).AsrtEq (WsNyAft)
End Sub

Private Sub Db__Tst()
Dim Db As Db: Set Db = Nw.Fx("N:\SapAccessReports\DutyPrepay5\SAPDownloadExcel\KE24 2010-01c.xls").Db
Ay(Db.Tny).Dmp
Db.Db.Close
End Sub

Sub Tst()
End Sub

Function FxOpn(Fx) As Workbook
Set FxOpn = Xls.Workbooks.Open(Fx)
End Function
