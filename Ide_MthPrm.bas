Attribute VB_Name = "Ide_MthPrm"
Option Compare Database
Option Explicit
Type PrmTy
    TyChr As String
    TyAsNm As String
    IsAy As Boolean
End Type
Type MthPrm
    Nm As String
    IsOpt As Boolean
    IsPrmAy As Boolean
    Ty As PrmTy
    DftVal As String
End Type
Type MthArg
    HasRetVal As Boolean
    PrmAy() As MthPrm
    RetTy As PrmTy
End Type

Function MthArg(MthLin$) As MthArg
Dim O As MthArg
With O
    .HasRetVal = MthLinHasRetVal(MthLin)
    .PrmAy = MthLinPrmAy(MthLin)
    .RetTy = MthLinRetTy(MthLin)
End With
MthArg = O
End Function

Function MthArgPrmNy(A As MthArg) As String()
Dim O$(), J%
For J = 0 To MthPrmUB(A.PrmAy)
    Push O, A.PrmAy(J).Nm
Next
MthArgPrmNy = O
End Function

Function MthLinArgStr$(MthLin$)
MthLinArgStr = TakBetBkt(MthLin)
End Function

Function MthLinHasRetVal(MthLin$ _
) As Boolean
Const CSub$ = "MthLinHasRetVal"
Dim A As SrcLinBrk
    A = SrcLinBrk(MthLin)
Select Case A.Ty
Case "Function", "Get": MthLinHasRetVal = True
Case "": Er CSub, "Give {MthLin} is not MthLin", MthLin
End Select
End Function

Function MthLinPrmAy(MthLin$) As MthPrm()
Dim ArgStr$
    ArgStr = TakBetBkt(MthLin, "()")
Dim P$()
    P = SplitComma(ArgStr)
Dim O() As MthPrm
    Dim U%: U = UB(P)
    ReDim O(U)
    Dim J%
    For J = 0 To U
        O(J) = MthPrm(P(J))
    Next
MthLinPrmAy = O
End Function

Function MthLinRetTy(MthLin$) As PrmTy
Dim L$
    L = MthLin
ParseSrcLinBrk L
With MthLinRetTy
    .TyChr = ParseTyChr(L)
    If .TyChr <> "" Then Exit Function
    L = TakAftBkt(L)
    If L = "" Then Exit Function
    If Not IsPfx(L, " As ") Then Stop
    L = RmvPfx(L, " As ")
    If IsSfx(L, "()") Then
        .IsAy = True
        L = RmvSfx(L, "()")
    End If
    .TyAsNm = L
    Exit Function
End With
End Function

Function MthPrm(MthPrmStr$) As MthPrm
Dim L$: L = MthPrmStr
Dim TyChr$
With MthPrm
    .IsOpt = ParseHasPfxSpc(L, "Optional")
    .IsPrmAy = ParseHasPfxSpc(L, "ParamArray")
    .Nm = ParseNm(L)
    .Ty.TyChr = ParseOneOfChr(L, "!@#$%^&")
    .Ty.IsAy = ParseHasPfx(L, "()")
End With
End Function

Sub MthPrmPush(O() As MthPrm, I As MthPrm)
Dim N&: N = MthPrmSz(O)
ReDim Preserve O(N)
O(N) = I
End Sub

Function MthPrmSz&(A() As MthPrm)
On Error Resume Next
MthPrmSz = UBound(A) + 1
End Function

Function MthPrmUB&(A() As MthPrm)
MthPrmUB = MthPrmSz(A) - 1
End Function

Function PrmTyAsTyNm$(A As PrmTy)
With A
    If .TyChr <> "" Then PrmTyAsTyNm = TyChrAsTyStr(.TyChr): Exit Function
    If .TyAsNm = "" Then
        PrmTyAsTyNm = "Variant"
    Else
        PrmTyAsTyNm = .TyAsNm
    End If
End With
End Function

Function TyChrAsTyStr$(TyChr$)
Dim O$
Select Case TyChr
Case "#": O = "Double"
Case "%": O = "Integer"
Case "!": O = "Signle"
Case "@": O = "Currency"
Case "^": O = "LongLong"
Case "$": O = "String"
Case "&": O = "Long"
Case Else: Stop
End Select
TyChrAsTyStr = O
End Function

Private Sub MthLinRetTy__Tst()
Dim MthLin$
Dim A As PrmTy:
MthLin = "Function MthPrm(MthPrmStr$) As MthPrm"
A = MthLinRetTy(MthLin)
Debug.Assert A.TyAsNm = "MthPrm"
Debug.Assert A.IsAy = False
Debug.Assert A.TyChr = ""

MthLin = "Function MthPrm(MthPrmStr$) As MthPrm()"
A = MthLinRetTy(MthLin)
Debug.Assert A.TyAsNm = "MthPrm"
Debug.Assert A.IsAy = True
Debug.Assert A.TyChr = ""

MthLin = "Function MthPrm$(MthPrmStr$)"
A = MthLinRetTy(MthLin)
Debug.Assert A.TyAsNm = ""
Debug.Assert A.IsAy = False
Debug.Assert A.TyChr = "$"

MthLin = "Function MthPrm(MthPrmStr$)"
A = MthLinRetTy(MthLin)
Debug.Assert A.TyAsNm = ""
Debug.Assert A.IsAy = False
Debug.Assert A.TyChr = ""
End Sub
