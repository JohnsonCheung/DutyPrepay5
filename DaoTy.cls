VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DaoTy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

Sub AsrtT(T, Optional D As Dao.Database)
On Error GoTo X:
Dim A$
A = D.TableDefs(T).Name
Exit Sub
X:
Err.Raise 1, , "Tbl[" & T & "] not found in Db[" & D.Name & "]"
End Sub





