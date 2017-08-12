VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tbl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Private A As Dao.TableDef
Function Init(Tbl As Dao.TableDef) As Tbl
Set A = Tbl
Set Init = Me
End Function
Function Rs() As Rs
Set Rs = ccNew.Rs(A.OpenRecordset)
End Function
Function Fny() As String()
Fny = Flds.Fny
End Function
Function Flds() As Flds
Set Flds = ccNew.Flds(A.Fields)
End Function

Function HasFld(F) As Boolean
HasFld = Flds.HasFld(F)
End Function

