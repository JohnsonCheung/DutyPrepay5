VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Db_InsDt_SqlAy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private Db As Db
Private Dt As Dt
Function SqlAy(Db_ As Db, Dt_ As Dt) As String()
Set Db = Db_
Set Dt = Dt_
If IsNoDta Then Exit Function
Dim mTp$
    mTp = Tp
Dim O$()
    Dim Dr
    ReDim O(Dt.URow)
    Dim J%
    J = 0
    For Each Dr In Dt.Dry.Dry
        O(J) = FmtQQAv(Tp, Dr)
        J = J + 1
    Next
SqlAy = O
End Function
Function SqlAyX(Db As Db, Dt As Dt) As Ay
Set SqlAyX = Nw.Ay(SqlAy(Db, Dt))
End Function
Function IsNoDta() As Boolean
IsNoDta = Dt.IsEmpty
End Function
Private Function TblNm$()
TblNm = Dt.DtNm
End Function
Private Function DbT() As DbT
Set DbT = Db.DbT(TblNm)
End Function
Private Function Tp$()
Tp = FmtQQ("Insert into [?] (?) values(?)", TblNm, FldLis, ValTp)
End Function
Private Function FldLis$()
FldLis = Dt.FldLis
End Function
Private Function ValTp$()
ValTp = DbT.InsSqlValPhraseTp
End Function

Private Sub ValTp__Tst()
Set Db = Nw.Db
Set Dt = Nw.Db.DbT("Permit").Dt
Debug.Print ValTp
End Sub

Private Sub FldLis__Tst()
Set Db = Nw.Db
Set Dt = Nw.Db.DbT("Permit").Dt
Debug.Print FldLis
End Sub

Private Sub Tp__Tst()
Set Db = Nw.Db
Set Dt = Nw.Db.DbT("Permit").Dt
Debug.Print Tp
End Sub

Sub Tst()
Tp__Tst
End Sub
Private Sub SqlAy__Tst()
Dim Db As Db: Set Db = Nw.Db
Dim Dt As Dt: Set Dt = Db.DbT("Permit").Dt
SqlAyX(Db, Dt).Brw
End Sub


