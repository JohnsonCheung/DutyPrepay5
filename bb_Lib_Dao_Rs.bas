Attribute VB_Name = "bb_Lib_Dao_Rs"
Option Compare Database
Option Explicit
Function RsDry(Rs As Dao.Recordset) As Variant()
Dim O()
With Rs
    While Not .EOF
        Push O, FldsDr(Rs.Fields)
        .MoveNext
    Wend
End With
RsDry = O
End Function
Function RsDrs(Rs As Dao.Recordset) As Drs

End Function
Function RsFny(Rs As Dao.Recordset) As String()
RsFny = FldsFny(Rs.Fields)
End Function
