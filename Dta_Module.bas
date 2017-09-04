Attribute VB_Name = "Dta_Module"
Option Explicit
Option Compare Database

Type Dt
    DtNm As String
    Fny() As String
    Dry() As Variant
End Type
Type Drs
    Fny() As String
    Dry() As Variant
End Type
Type Ds
    DsNm As String
    DtAy() As Dt
End Type

Function AySel(Ay, IdxAy&())
Dim U&
    U = UB(IdxAy)
Dim O
    O = Ay
    ReDim O(U)
Dim J&
For J = 0 To U
    O(J) = Ay(IdxAy(J))
Next
AySel = O
End Function

Function DbDs(TblNmLvs_or_Tny, Optional DsNm$ = "Ds", Optional D As Database) As Ds
Dim Tny$(): Tny = NyCv(TblNmLvs_or_Tny)
Dim O As Ds
    O.DsNm = DsNm
    Dim J%
    Dim U%: U = UB(Tny)
    ReDim O.DtAy(U)
    For J = 0 To UB(Tny)
        O.DtAy(J) = TblDt(Tny(J), D)
    Next
DbDs = O
End Function

Function DrExpLinesCol(Dr, LinesColIdx%) As Variant()
Dim Ay$()
    Ay = SplitCrLf(Dr(LinesColIdx))
Dim O()
    Dim IDr
        IDr = Dr
    Dim I
    For Each I In Ay
        IDr(LinesColIdx) = I
        Push O, IDr
    Next
DrExpLinesCol = O
End Function

Function DrsAddRowIdxCol(A As Drs) As Drs
Dim O As Drs
    O.Fny = A.Fny
    AyIns O.Fny, "RowIdx"
Dim ODry()
    If Not AyIsEmpty(A.Dry) Then
        Dim J&, Dr
        For Each Dr In A.Dry
            AyIns Dr, J: J = J + 1
            Push ODry, Dr
        Next
    End If
O.Dry = ODry
DrsAddRowIdxCol = O
End Function

Sub DrsBrw(Drs As Drs, Optional MaxColWdt& = 100, Optional BrkColNm$)
AyBrw DrsLy(Drs, MaxColWdt, BrkColNm$)
End Sub

Function DrsCol(Drs As Drs, ColNm$) As Variant()
Dim ColIdx%: ColIdx = AyIdx(Drs.Fny, ColNm)
DrsCol = DryCol(Drs.Dry, ColIdx)
End Function

Function DrsExpLinesCol(Drs As Drs, LinesColNm$) As Drs
Dim Idx%
    Idx = AyIdx(Drs.Fny, LinesColNm)
Dim Dry()
    Dim Dr
    For Each Dr In Drs.Dry
        PushAy Dry, DrExpLinesCol(Dr, Idx)
    Next
Dim O As Drs
    O.Fny = Drs.Fny
    O.Dry = Dry
DrsExpLinesCol = O
End Function

Function DrsSel(A As Drs, Fny) As Drs
Dim mFny$(): mFny = NyCv(Fny)
Dim IdxAy&()
    IdxAy = AyIdxAy(A.Fny, mFny)
Dim Dry()
    Dim Dr
    For Each Dr In A.Dry
        Push Dry, AySel(Dr, IdxAy)
    Next
Dim O As Drs
    O.Fny = mFny
    O.Dry = Dry
DrsSel = O
End Function

Function DrsStrCol(Drs As Drs, ColNm$) As String()
DrsStrCol = AySy(DrsCol(Drs, ColNm))
End Function

Sub DryBrw(Dry)
AyBrw DryLy(Dry)
End Sub

Sub DsBrw(A As Ds)
AyBrw DsLy(A)
End Sub

Function DsLy(A As Ds, Optional MaxColWdt& = 1000, Optional BrkLinMapStr$) As String()
Dim O$()
    Push O, "*Ds " & A.DsNm
Dim Dic As Dictionary ' DicOf_TblNm_to_BrkColNm
    Set Dic = MapDic(BrkMapStr(BrkLinMapStr))
If Not IsEmptyDtAy(A.DtAy) Then
    Dim J%, DtNm$, Dt As Dt, BrkColNm$
    For J = 0 To UBound(A.DtAy)
        Dt = A.DtAy(J)
        DtNm$ = Dt.DtNm
        If Dic.Exists(DtNm) Then BrkColNm = Dic(DtNm) Else BrkColNm = ""
        PushAy O, DtLy(Dt, MaxColWdt, BrkColNm)
    Next
End If
DsLy = O
End Function

Function DsNew(TblNmLvs$, Optional DsNm$ = "Ds", Optional D As Database) As Ds
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
DsNew = O
End Function

Function DtAySz%(DtAy() As Dt)
On Error Resume Next
DtAySz = UBound(DtAy) + 1
End Function

Sub DtBrw(Dt As Dt)
AyBrw DtLy(Dt)
End Sub

Function DtCsvLy(A As Dt) As String()
Dim O$()
Dim QQStr$
Dim Dr
Push O, JnComma(DblAyQuote(A.Fny))
For Each Dr In A.Dry
    Push O, FmtQQAv(QQStr, Dr)
Next
End Function

Sub DtDmp(Dt As Dt)
AyDmp DtLy(Dt)
End Sub

Function DtDrs(A As Dt) As Drs
Dim O As Drs
O.Fny = A.Fny
O.Dry = A.Dry
DtDrs = O
End Function

Sub Fiy(Fny$(), FldNmLvs$, ParamArray OAp())
'Fiy=Field Index Array
Dim Ay$(): Ay = SplitSpc(FldNmLvs)
Dim I&(): I = AyIdxAy(Fny, Ay)
Dim J%
For J = 0 To UB(I)
    OAp(J) = I(J)
Next
End Sub

Function FnyLIdxAy(Fny$(), FldNmLvs$) As Long()
'FnyL = Fny + FldNmLvs
'Return Field Idx Ay
FnyLIdxAy = AyIdxAy(Fny, SplitSpc(FldNmLvs))
End Function

Function IsEmptyDs(A As Ds) As Boolean
IsEmptyDs = IsEmptyDtAy(A.DtAy)
End Function

Function IsEmptyDt(A As Dt) As Boolean
IsEmptyDt = AyIsEmpty(A.Dry)
End Function

Function IsEmptyDtAy(DtAy() As Dt) As Boolean
IsEmptyDtAy = DtAySz(DtAy) = 0
End Function

Function NyCv(Ny) As String()
If IsStrAy(Ny) Then NyCv = Ny: Exit Function
If Not IsStr(Ny) Then Err.Raise 1, , "NyCv: Given [Ny] must be StrAy or Str, but now [" & TypeName(Ny) & "]"
NyCv = SplitLvs(Ny)
End Function

Function ObjCollSel(ObjColl, PrpNy) As Drs
Dim Ny$()
    Ny = NyCv(PrpNy)
Dim Dry()
    Dim Obj
    If Not IsEmptyColl(ObjColl) Then
        For Each Obj In ObjColl
            Push Dry, ObjSelPrp(Obj, Ny)
        Next
    End If
Dim O As Drs
    O.Fny = Ny
    O.Dry = Dry
ObjCollSel = O
End Function

Function ObjSelPrp(Obj, PrpNy$()) As Variant()
Dim U%
    U = UB(PrpNy)
Dim O()
    ReDim O(U)
    Dim J%
    For J = 0 To U
        O(J) = CallByName(Obj, PrpNy(J), VbGet)
    Next
ObjSelPrp = O
End Function

Function TblDt(T, Optional D As Database) As Dt
Dim O As Dt
O.DtNm = T
O.Dry = RsDry(Tbl(T, D).OpenRecordset)
O.Fny = TblFny(T, D)
TblDt = O
End Function

Private Sub DrsSel__Tst()
DrsBrw DrsSel(MdFunDrs, "MdNm FunNm Mdy Ty")
End Sub

Private Sub DsNew__Tst()
Dim Ds As Ds
Ds = DsNew("Permit PermitD")
Stop
End Sub

Private Sub ObjCollSel__Tst()
DrsBrw ObjCollSel(Flds("Permit"), "Name Type")
Stop
End Sub

Sub Tst()
DrsSel__Tst
DsNew__Tst
ObjCollSel__Tst
End Sub
