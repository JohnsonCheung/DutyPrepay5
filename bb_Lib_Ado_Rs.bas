Attribute VB_Name = "bb_Lib_Ado_Rs"
Option Compare Database
Option Explicit
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
Function ARsDrs(Rs As ADODB.Recordset) As Drs
Dim O As Drs
O.Fny = ARsFny(Rs)
O.Dry = ARsDry(Rs)
ARsDrs = O
End Function
Function ARsFny(Rs As ADODB.Recordset) As String()
ARsFny = AFldsFny(Rs.Fields)
End Function

