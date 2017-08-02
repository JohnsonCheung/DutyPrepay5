Attribute VB_Name = "bb_Lib_Vb_Str"
Option Compare Database
Option Explicit
Function HasSubStr(S, SubStr) As Boolean
HasSubStr = InStr(S, SubStr) > 0
End Function
Function Quote$(S, QuoteStr$)
With BrkQuote(QuoteStr)
    Quote = .S1 & S & .S2
End With
End Function
Sub TakBet__Tst()
Const S1$ = "Excel 8.0;HDR=YES;IMEX=2;DATABASE=??"
Const S2$ = "Excel 8.0;HDR=YES;IMEX=2;DATABASE=??;AA=XX"
Debug.Assert TakBet(S1, "DATABASE=", ";") = "??"
Debug.Assert TakBet(S2, "DATABASE=", ";") = "??"
End Sub
Function TakBet$(S, S1, S2, Optional NoTrim As Boolean)
With Brk1(S, S1, NoTrim)
    If .S2 = "" Then Exit Function
    TakBet = Brk1(.S2, S2, NoTrim).S1
End With
End Function
Function FstChr$(S)
FstChr = Left(S, 1)
End Function
Function LasChr$(S)
LasChr = Right(S, 1)
End Function
Function AlignL$(S, W)
Dim L%:
If IsNull(S) Then
    L = 0
Else
    L = Len(S)
End If
If W > L Then
    AlignL = S & Space(W - L)
Else
    AlignL = S
End If
End Function
Function WrtStr(S, Ft)
Dim T As TextStream
Set T = Fso.OpenTextFile(Ft, ForWriting, True)
T.Write S
T.Close
End Function
Function AlignR$(S, W%)
Dim L%: L = Len(S)
If W > L Then
    AlignR = Space(W - L) & S
Else
    AlignR = S
End If
End Function

