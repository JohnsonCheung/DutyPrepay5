Attribute VB_Name = "bb_Lib_Vb_Str"
Option Compare Database
Option Explicit

Function AlignL$(S, W)
Dim L%:
If IsNull(S) Then
    L = 0
Else
    L = Len(S)
End If
If W >= L Then
    AlignL = S & Space(W - L)
Else
    If W > 2 Then
        AlignL = Left(S, W - 2) + ".."
    Else
        AlignL = Left(S, W)
    End If
End If
End Function

Function AlignR$(S, W%)
Dim L%: L = Len(S)
If W > L Then
    AlignR = Space(W - L) & S
Else
    AlignR = S
End If
End Function

Function FstChr$(S)
FstChr = Left(S, 1)
End Function

Function HasSubStr(S, SubStr) As Boolean
HasSubStr = InStr(S, SubStr) > 0
End Function

Function IsDigit(C) As Boolean
IsDigit = "0" <= C And C <= "9"
End Function

Function IsLetter(C) As Boolean
Dim C1$: C1 = UCase(C)
IsLetter = ("A" <= C1 And C1 <= "Z")
End Function

Function IsNmChr(C) As Boolean
IsNmChr = True
If IsLetter(C) Then Exit Function
If C = "_" Then Exit Function
If IsDigit(C) Then Exit Function
IsNmChr = False
End Function

Function IsSfx(S, Sfx) As Boolean
IsSfx = (Right(S, Len(Sfx)) = Sfx)
End Function

Function LasChr$(S)
LasChr = Right(S, 1)
End Function

Function LasLin$(S)
LasLin = AyLasEle(SplitCrLf(S))
End Function

Function LinesCnt&(Lines$)
LinesCnt = Sz(SplitCrLf(Lines))
End Function

Function Quote$(S, QuoteStr$)
With BrkQuote(QuoteStr)
    Quote = .S1 & CStr(S) & .S2
End With
End Function

Sub StrBrw(S)
Dim T$: T = TmpFt
StrWrt S, T
FtBrw T
End Sub

Function StrWrt(S, Ft)
Dim F%: F = FreeFile(1)
Open Ft For Output As #F
Print #F, S
Close #F
'Dim T As TextStream
'Set T = Fso.OpenTextFile(Ft, ForWriting, True)
'T.Write S
'T.Close
End Function

Function TakBet$(S, S1, S2, Optional NoTrim As Boolean)
With Brk1(S, S1, NoTrim)
    If .S2 = "" Then Exit Function
    TakBet = Brk1(.S2, S2, NoTrim).S1
End With
End Function

Private Sub TakBet__Tst()
Const S1$ = "Excel 8.0;HDR=YES;IMEX=2;DATABASE=??"
Const S2$ = "Excel 8.0;HDR=YES;IMEX=2;DATABASE=??;AA=XX"
Debug.Assert TakBet(S1, "DATABASE=", ";") = "??"
Debug.Assert TakBet(S2, "DATABASE=", ";") = "??"
End Sub

Sub Tst()
TakBet__Tst
End Sub
