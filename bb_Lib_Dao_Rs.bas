Attribute VB_Name = "bb_Lib_Dao_Rs"
Option Compare Database
Option Explicit

Function RsDrs(A As DAO.Recordset) As Drs
RsDrs.Fny = RsFny(A)
RsDrs.Dry = RsDry(A)
End Function

Function RsDry(A As DAO.Recordset) As Variant()
Dim O()
With A
    While Not .EOF
        Push O, FldsDr(A.Fields)
        .MoveNext
    Wend
End With
RsDry = O
End Function

Function RsFny(A As DAO.Recordset) As String()
RsFny = FldsFny(A.Fields)
End Function

Function RsSy(A As DAO.Recordset) As String()
Dim O$()
With A
    While Not .EOF
        Push O$, A.Fields(0).Value
        .MoveNext
    Wend
End With
RsSy = O
End Function
