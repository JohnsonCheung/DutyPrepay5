Attribute VB_Name = "Ide_SrcLin"
Option Compare Database
Option Explicit


Type SrcLinBrk
    MthNm As String
    Ty As String    ' Sub Function Get Set Let (Ty here means SrcTy)
    Mdy As String
End Type

Sub AssertIsMthLin(SrcLin)
If Not IsMthLin(SrcLin) Then Er "AssertIsMthLin", "{SrcLin} is not MthLin", SrcLin
End Sub

Function IsMthLin(SrcLin) As Boolean
IsMthLin = SrcLinMthNm(SrcLin) <> ""
End Function

Function MthLinEndLinPfx$(MthLin)
MthLinEndLinPfx = "End " & SrcLinMthTy(MthLin)
End Function

Function MthLinMthTy$(MthLin)
AssertIsMthLin MthLin
Dim O$:
    O = SrcLinMthTy(MthLin)
If O = "" Then Er "MthLinMthTy", "{MthLin} has been asserted to be MthLin, but SrcLinMthTy(MthLin) gives empty string", MthLin
MthLinMthTy = O
End Function

Function MthLinRmvMdy$(MthLin)
AssertIsMthLin MthLin
Dim O$: O = MthLin
ParseMdy O
MthLinRmvMdy = O
End Function

Function MthLinRplMdy$(MthLin, ToMdy)
AssertIsMdy ToMdy
AssertIsMthLin MthLin
MthLinRplMdy = StrAppSpc(ToMdy) & MthLinRmvMdy(MthLin)
End Function

Function SrcLinBrk(SrcLin) As SrcLinBrk
Dim L$
    L = SrcLin
SrcLinBrk = ParseSrcLinBrk(L)
End Function

Function SrcLinDr(SrcLin, Lno&) As Variant()
With SrcLinBrk(SrcLin)
    SrcLinDr = Array(Lno, .Mdy, .Ty, .MthNm)
End With
End Function

Function SrcLinIsCd(Lin) As Boolean
Dim L$: L = Trim(Lin)
If L = "" Then Exit Function
If FstChr(L) = "'" Then Exit Function
SrcLinIsCd = True
End Function

Function SrcLinIsEnm(SrcLin) As Boolean
Dim L$: L = SrcLin
ParseMdy L
SrcLinIsEnm = IsPfx(L, "Enum")
End Function

Function SrcLinIsRmk(SrcLin) As Boolean
SrcLinIsRmk = FstChr(LTrim(SrcLin)) = "'"
End Function

Function SrcLinIsTy(SrcLin) As Boolean
Dim L$: L = SrcLin
ParseMdy L
SrcLinIsTy = IsPfx(L, "Type")
End Function

Function SrcLinMthNm$(SrcLin)
Dim L$: L = SrcLin
ParseMdy L
Dim MthTy$
    MthTy = ParseMthTy(L)
If MthTy = "" Then Exit Function
If MthTy = "Property" Then
    If ParsePrpTy(L) = "" Then Stop
End If
SrcLinMthNm = ParseNm(L)
End Function

Function SrcLinMthTy$(SrcLin)
Dim L$: L = Trim(SrcLin)
ParseMdy L
SrcLinMthTy = ParseMthTy(L)
End Function
