Attribute VB_Name = "bb_Lib_Dta"
Option Compare Database
Option Explicit
Type Drs
    Fny() As String
    Dry() As Variant
End Type
Type Dt
    DtNm As String
    Fny() As String
    Dry() As Variant
End Type
Type Ds
    DsNm As String
    DtAy() As Dt
End Type
Function DrsCol(Drs As Drs, ColNm$) As Variant()
Dim ColIdx%: ColIdx = AyIdx(Drs.Fny, ColNm)
DrsCol = DryCol(Drs.Dry, ColIdx)
End Function
Function DrsStrCol(Drs As Drs, ColNm$) As String()
DrsStrCol = AySy(DrsCol(Drs, ColNm))
End Function
Function DtCsvLy(A As Dt) As String()
Dim O$()
Dim QQStr$
Dim Dr
Push O, JnComma(DblQuoteAy(A.Fny))
For Each Dr In A.Dry
    Push O, FmtQQAv(QQStr, Dr)
Next
End Function
Function IsEmptyDt(A As Dt) As Boolean
IsEmptyDt = IsEmptyAy(A.Dry)
End Function
Function IsEmptyDs(A As Ds) As Boolean
IsEmptyDs = IsEmptyDtAy(A.DtAy)
End Function
Function IsEmptyDtAy(DtAy() As Dt) As Boolean
IsEmptyDtAy = DtAySz(DtAy) = 0
End Function
Function DtAySz%(DtAy() As Dt)
On Error Resume Next
DtAySz = UBound(DtAy) + 1
End Function
Sub BrwDt(Dt As Dt)
BrwAy DtLy(Dt)
End Sub
Sub BrwDry(Dry)
BrwAy DryLy(Dry)
End Sub
Sub BrwDrs(Drs As Drs)
BrwAy DrsLy(Drs)
End Sub
Sub NewDs__Tst()
Dim Ds As Ds
Ds = NewDs("Permit PermitD")
Stop
End Sub
Sub DmpDt(Dt As Dt)
DmpAy DtLy(Dt)
End Sub
Sub SelObjColl__Tst()
BrwDrs SelObjColl(Flds("Permit"), "Name Type")
Stop
End Sub
Function SelObjColl(ObjColl, PrpNy) As Drs
Dim Ny$()
    Ny = CvNy(PrpNy)
Dim Dry()
    Dim Obj
    If Not IsEmptyColl(ObjColl) Then
        For Each Obj In ObjColl
            Push Dry, SelObjPrp(Obj, Ny)
        Next
    End If
Dim O As Drs
    O.Fny = Ny
    O.Dry = Dry
SelObjColl = O
End Function
Function SelObjPrp(Obj, PrpNy$()) As Variant()
Dim U%
    U = UB(PrpNy)
Dim O()
    ReDim O(U)
    Dim J%
    For J = 0 To U
        O(J) = CallByName(Obj, PrpNy(J), VbGet)
    Next
SelObjPrp = O
End Function
Function DtLy(Dt As Dt) As String()
Dim Rs As Drs
    Rs.Fny = Dt.Fny
    Rs.Dry = Dt.Dry
Dim O$()
    Push O, "*Tbl " & Dt.DtNm
    PushAy O, DrsLy(Rs)
DtLy = O
End Function
Function CvNy(Ny) As String()
If IsStrAy(Ny) Then CvNy = Ny: Exit Function
If Not IsStr(Ny) Then Err.Raise 1, , "CvNy: Given [Ny] must be StrAy or Str, but now [" & TypeName(Ny) & "]"
CvNy = SplitLvs(Ny)
End Function
Function SelAy(Ay, IdxAy&())
Dim U&
    U = UB(IdxAy)
Dim O
    O = Ay
    ReDim O(U)
Dim J&
For J = 0 To U
    O(J) = Ay(IdxAy(J))
Next
SelAy = O
End Function
Sub SelDrs__Tst()
BrwDrs SelDrs(MdFunDrs, "MdNm FunNm Mdy Ty")
End Sub
Function SelDrs(A As Drs, Fny) As Drs
Dim mFny$(): mFny = CvNy(Fny)
Dim IdxAy&()
    IdxAy = AyIdxAy(A.Fny, mFny)
Dim Dry()
    Dim Dr
    For Each Dr In A.Dry
        Push Dry, SelAy(Dr, IdxAy)
    Next
Dim O As Drs
    O.Fny = mFny
    O.Dry = Dry
SelDrs = O
End Function
Function DrsLy(Drs As Drs) As String()
If IsEmptyAy(Drs.Fny) Then Exit Function
Dim Dry(): Dry = Drs.Dry
Push Dry, Drs.Fny
Dim Ay$(): Ay = DryLy(Dry)
Dim Lin$: Lin = Pop(Ay)
Dim Hdr$: Hdr = Pop(Ay)
Dim O$()
    PushAy O, Array(Lin, Hdr)
    PushAy O, Ay
    Push O, Lin
DrsLy = O
End Function
Function TblDt(T, Optional D As Database) As Dt
Dim O As Dt
O.DtNm = T
O.Dry = RsDry(Tbl(T, D).OpenRecordset)
O.Fny = TblFny(T, D)
TblDt = O
End Function
Function NewDs(TblNmLvs$, Optional DsNm$ = "Ds", Optional D As Database) As Ds
Dim DtAy() As Dt
    Dim U%, Tny$()
    Tny = SplitLvs(TblNmLvs)
    U = UB(Tny)
    ReDim DtAy(U)
    Dim J%
    For J = 0 To U
        DtAy(J) = TblDt(Tny(J), D)
    Next
Dim O As Ds
    O.DsNm = DsNm
    O.DtAy = DtAy
NewDs = O
End Function

