Attribute VB_Name = "Dao_Rs"
Option Explicit
Option Compare Database

Function RsDrs(A As Dao.Recordset) As Drs
RsDrs.Fny = RsFny(A)
RsDrs.Dry = RsDry(A)
End Function

Function RsDry(A As Dao.Recordset) As Variant()
Dim O()
With A
    While Not .EOF
        Push O, FldsDr(A.Fields)
        .MoveNext
    Wend
End With
RsDry = O
End Function

Function RsFny(A As Dao.Recordset) As String()
RsFny = FldsFny(A.Fields)
End Function

Function RsSy(A As Dao.Recordset) As String()
Dim O$()
With A
    While Not .EOF
        Push O$, A.Fields(0).Value
        .MoveNext
    Wend
End With
RsSy = O
End Function
