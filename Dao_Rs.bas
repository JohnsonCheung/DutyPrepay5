Attribute VB_Name = "Dao_Rs"
Option Explicit
Option Compare Database

Function RsDrs(a As Dao.Recordset) As Drs
RsDrs.Fny = RsFny(a)
RsDrs.Dry = RsDry(a)
End Function

Function RsDry(a As Dao.Recordset) As Variant()
Dim O()
With a
    While Not .EOF
        Push O, FldsDr(a.Fields)
        .MoveNext
    Wend
End With
RsDry = O
End Function

Function RsFny(a As Dao.Recordset) As String()
RsFny = FldsFny(a.Fields)
End Function

Function RsSy(a As Dao.Recordset) As String()
Dim O$()
With a
    While Not .EOF
        Push O$, a.Fields(0).Value
        .MoveNext
    Wend
End With
RsSy = O
End Function
