Attribute VB_Name = "Ide_Parse"
Option Compare Database
Option Explicit

Function ParseConst$(OLin$, ConstStr$)
If IsPfx(OLin, ConstStr) Then
    ParseConst = ConstStr
    OLin = RmvPfx(OLin, ConstStr)
End If
End Function

Function ParseConstSpc$(OLin$, ConstStr$)
ParseConstSpc = ParseConst(OLin, ConstStr & " ")
End Function

Function ParseHasPfx(OLin$, Pfx$) As Boolean
ParseHasPfx = ParseConst(OLin, Pfx) = Pfx
End Function

Function ParseHasPfxSpc(OLin$, Pfx$) As Boolean
ParseHasPfxSpc = ParseHasPfx(OLin, Pfx & " ")
End Function

Function ParseMdy$(OLin$)
ParseMdy = RTrim(ParseOneOf(OLin, SyOfMdy))
End Function

Function ParseMthTy$(OLin$)
ParseMthTy = RTrim(ParseOneOf(OLin, SyOfMthTy))
End Function

Function ParseNm$(OLin$)
Dim J%
J = 1
If Not IsLetter(FstChr(OLin)) Then GoTo Nxt
For J = 2 To Len(OLin)
    If Not IsNmChr(Mid(OLin, J, 1)) Then GoTo Nxt
Next
Nxt:
If J = 1 Then Exit Function
ParseNm = Left(OLin, J - 1)
OLin = Mid(OLin, J)
End Function

Function ParseOneOf(OLin$, OneOfAy$())
Dim I
For Each I In OneOfAy
    If IsPfx(OLin, I) Then OLin = RmvPfx(OLin, I): ParseOneOf = I: Exit Function
Next
End Function

Function ParseOneOfChr$(OLin$, LisOfChr$)
Dim C$: C = FstChr(OLin)
If HasSubStr(LisOfChr, C) Then
    OLin = RmvFstChr(OLin)
    ParseOneOfChr = C
End If
End Function

Function ParsePrpTy$(OLin$)
ParsePrpTy = RTrim(ParseOneOf(OLin, SyOfPrpTy))
End Function

Function ParseSrcLinBrk(OLin$) As SrcLinBrk
Const CSub$ = "ParseSrcLinBrk"
Dim O As SrcLinBrk
O.Mdy = ParseMdy(OLin)
O.Ty = ParseMthTy(OLin)
Select Case O.Ty
Case "Property": O.Ty = ParsePrpTy(OLin)
Case "": Exit Function
End Select
If O.Ty = "" Then Stop
O.MthNm = ParseNm(OLin)
If O.MthNm = "" Then Er CSub, "{OLin} does not have a [function name]"
ParseSrcLinBrk = O
End Function

Function ParseTyChr$(OLin$)
Dim Fst$: Fst = FstChr(OLin)
If HasSubStr("!@#$%^&", Fst) Then
    ParseTyChr = Fst
    OLin = RmvFstChr(OLin)
End If
End Function

Private Function SyOfDfnTy() As String()
Static X As Boolean, Y
If Not X Then
    X = True
    Y = Sy("Function ", "Sub ", "Property ", "Type ", "Enum ")
End If
SyOfDfnTy = Y
End Function

Private Function SyOfMdy() As String()
Static X As Boolean, Y
If Not X Then
    X = True
    Y = Sy("Public ", "Private ", "Friend ")
End If
SyOfMdy = Y
End Function

Private Function SyOfMthTy() As String()
Static X As Boolean, Y
If Not X Then
    X = True
    Y = Sy("Function ", "Sub ", "Property ")
End If
SyOfMthTy = Y
End Function

Private Function SyOfPrpTy() As String()
Static X As Boolean, Y
If Not X Then
    X = True
    Y = Sy("Get ", "Set ", "Let ")
End If
SyOfPrpTy = Y
End Function
