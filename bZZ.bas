Attribute VB_Name = "bZZ"
Option Compare Database
Option Explicit
Sub SetLatestYYMMDD(ByRef oYY As Byte, ByRef oMM As Byte, ByRef oDD As Byte)
Dim M$: M = SqlToV("Select Max(YY*10000+MM*100+DD) from OH")
oYY = Left(M, 2)
oMM = Mid(M, 3, 2)
oDD = Right(M, 2)
End Sub
Function SqlToInt%(pSql$)
With CurrentDb.OpenRecordset(pSql)
    SqlToInt = .Fields(0).Value
    .Close
End With
End Function
Function SqlToV(pSql$)
With CurrentDb.OpenRecordset(pSql)
    SqlToV = .Fields(0).Value
    .Close
End With
End Function
Function SqlToAys(pSql$) As String()
Dim O$()
With CurrentDb.OpenRecordset(pSql)
    While Not .EOF
        If Not IsNull(.Fields(0).Value) Then Push O, .Fields(0).Value
        .MoveNext
    Wend
    .Close
End With
SqlToAys = O
End Function
