Attribute VB_Name = "bb_Lib_Dao"
Option Compare Database
Option Explicit

Sub AA1()
DbInfBrw
End Sub

Sub AddFld(T, F, Ty As Dao.DataTypeEnum, Optional D As Dao.Database)
Dim mFld As New Dao.Field
mFld.Name = F
mFld.Type = Ty
Flds(T, D).Append mFld
End Sub

Sub AssertT(T, Optional D As Dao.Database)
On Error GoTo X:
Dim A$
A = D.TableDefs(T).Name
Exit Sub
X:
Err.Raise 1, , "Tbl[" & T & "] not found in Db[" & D.Name & "]"
End Sub

Sub BrwSql(Sql$, Optional D As Dao.Database)
DrsBrw SqlDrs(Sql, D)
End Sub

Sub BrwTbl(T, Optional D As Dao.Database)
DtBrw TblDt(T, D)
End Sub

Function CurDb() As Database
Static X As Database
If IsNothing(X) Then Set X = CurrentDb
Set CurDb = X
End Function

Function DaoTyStr$(T As Dao.DataTypeEnum)
Dim O$
Select Case T
Case Dao.DataTypeEnum.dbBoolean: O = "Boolean"
Case Dao.DataTypeEnum.dbDouble: O = "Double"
Case Dao.DataTypeEnum.dbText: O = "Text"
Case Dao.DataTypeEnum.dbDate: O = "Date"
Case Dao.DataTypeEnum.dbByte: O = "Byte"
Case Dao.DataTypeEnum.dbInteger: O = "Int"
Case Dao.DataTypeEnum.dbLong: O = "Long"
Case Dao.DataTypeEnum.dbDouble: O = "Doubld"
Case Dao.DataTypeEnum.dbDate: O = "Date"
Case Dao.DataTypeEnum.dbDecimal: O = "Decimal"
Case Dao.DataTypeEnum.dbCurrency: O = "Currency"
Case Dao.DataTypeEnum.dbSingle: O = "Single"

Case Else: Stop
End Select
DaoTyStr = O
End Function

Sub DbInfBrw(Optional A As Database)
AyBrw DsLy(DbInfDs(A), 2000)
Exit Sub
WbVis DsWb(DbInfDs(A))
End Sub

Function DbInfDs(Optional A As Database) As Ds
Dim O As Ds
DsAddDt O, DbLnkInfDt(A)
DsAddDt O, DbStruInfDt(A)
DsAddDt O, DbTblFldInfDt(A)
O.DsNm = DftDb(A).Name
DbInfDs = O
End Function

Function DbLnkInfDt(Optional D As Database) As Dt
Dim T, Dry(), C$
For Each T In DbTny(D)
    C = Tbl(T, D).Connect
    If C <> "" Then Push Dry, Array(T, C)
Next
Dim O As Dt
O.Dry = Dry
O.Fny = Sy("Tbl", "Connect")
O.DtNm = "Lnk"
DbLnkInfDt = O
End Function

Function DbStruInfDt(Optional D As Database) As Dt
Dim T, Dry()
For Each T In DbTny(D)
    Push Dry, Array(T, TblRecCnt(T, D), TblDes(T, D), TblStruLin(T, SkipTblNm:=True, D:=D))
Next
Dim O As Dt
    With O
        .Dry = Dry
        .Fny = Sy("Tbl", "RecCnt", "Des", "Stru")
        .DtNm = "Tbl"
    End With
DbStruInfDt = O
End Function

Function DbTblFldInfDt(Optional D As Database) As Dt
Dim T, Dry()
For Each T In DbTny(D)
    PushAy Dry, TblFldInfDry(T, D)
Next
Dim O As Dt
O.Dry = Dry
O.Fny = TblFldInfFny
O.DtNm = "TblFld"
DbTblFldInfDt = O
End Function

Function DbTny(Optional A As Database) As String()
DbTny = SqlSy("Select Name from MSysObjects where Type in (1,6) and Left(Name,4)<>'MSYS'", A)
End Function

Function DftDb(D As Database) As Database
If IsNothing(D) Then
    Set DftDb = CurDb
Else
    Set DftDb = D
End If
End Function

Sub DrpTbl(T, Optional D As Database)
If IsTbl(T, D) Then D.Execute FmtQQ("Drop Table [?]", T)
End Sub

Sub DsAddDt(ODs As Ds, T As Dt)
If DsHasDt(ODs, T.DtNm) Then Err.Raise 1, , FmtQQ("DsAddDt: Ds[?] already has Dt[?]", ODs.DsNm, T.DtNm)
Dim N%: N = DtAySz(ODs.DtAy)
ReDim Preserve ODs.DtAy(N)
ODs.DtAy(N) = T
End Sub

Function DsHasDt(A As Ds, DtNm) As Boolean
If DsIsEmpty(A) Then Exit Function
Dim J%
For J = 0 To UBound(A.DtAy)
    If A.DtAy(J).DtNm = DtNm Then DsHasDt = True: Exit Function
Next
End Function

Function DsIsEmpty(A As Ds) As Boolean
DsIsEmpty = DtAySz(A.DtAy) = 0
End Function

Function Fld(T, F, Optional D As Database) As Dao.Field
Set Fld = Tbl(T, D).Fields(F)
End Function

Function FldDes$(F As Dao.Field)
FldDes = PrpVal(F.Properties, "Description")
End Function

Function FldInfDr(T, F, Optional D As Dao.Field) As Variant()
Dim FF As Dao.Field: Set FF = Fld(T, F, D)
With FF
    FldInfDr = Array(F, IIf(FldIsPk(T, F, D), "*", ""), DaoTyStr(.Type), .Size, .DefaultValue, .Required, FldDes(FF))
End With
End Function

Function FldInfFny() As String()
FldInfFny = SplitSpc("Fld Pk Ty Sz Dft Req Des")
End Function

Function FldIsPk(T, F, D As Database) As Boolean
FldIsPk = AyHas(TblPk(T, D), F)
End Function

Property Get Flds(T, Optional D As Dao.Database) As Dao.Fields
Set Flds = Tbl(T, D).Fields
End Property

Function FldsDr(Flds As Dao.Fields) As Variant()
Dim O()
ReDim O(Flds.Count - 1)
Dim J%, F As Dao.Field
For Each F In Flds
    O(J) = F.Value
    J = J + 1
Next
FldsDr = O
End Function

Function FldsFny(Flds As Dao.Fields) As String()
Dim O$()
Dim F As Dao.Field
For Each F In Flds
    Push O, F.Name
Next
FldsFny = O
End Function

Function FnyQuote(Fny$(), ToQuoteFny$()) As String()
If AyIsEmpty(Fny) Then Exit Function
Dim O$(): O = Fny
Dim J%, F
For Each F In O
    If AyHas(ToQuoteFny, F) Then O(J) = Quote(F, "[]")
    J = J + 1
Next
FnyQuote = O
End Function

Function FnyQuoteIfNeed(Fny$()) As String()
If AyIsEmpty(Fny) Then Exit Function
Dim O$(), J%, F
O = Fny
For Each F In Fny
    If IsNeedQuote(F) Then O(J) = Quote(F, "'")
    J = J + 1
Next
FnyQuoteIfNeed = O
End Function

Function HasFld(T, F, Optional D As Database) As Boolean
AssertT T, D
HasFld = HasFld_Tbl(Tbl(T, D), F)
End Function

Function HasFld_Flds(Flds As Dao.Fields, F) As Boolean
Dim I As Dao.Field
For Each I In Flds
    If I.Name = F Then HasFld_Flds = True: Exit Function
Next
End Function

Function HasFld_Tbl(T As Dao.TableDef, F) As Boolean
HasFld_Tbl = HasFld_Flds(T.Fields, F)
End Function

Function IsNeedQuote(S) As Boolean
IsNeedQuote = True
If HasSubStr(S, " ") Then Exit Function
If HasSubStr(S, "#") Then Exit Function
If HasSubStr(S, ".") Then Exit Function
IsNeedQuote = False
End Function

Function PrpVal(A As Dao.Properties, PrpNm)
On Error Resume Next
PrpVal = A(PrpNm).Value
End Function

Function Tbl(T, Optional D As Database) As Dao.TableDef
Set Tbl = DftDb(D).TableDefs(T)
End Function

Function TblDes$(T, D As Database)
TblDes = PrpVal(Tbl(T, D).Properties, "Description")
End Function

Function TblFld(T, F, Optional D As Database) As Dao.Field
Set TblFld = Tbl(T, D).Fields(F)
End Function

Function TblFldInfDry(T, D As Database) As Variant()
Dim O(), F, Dr(), Fny$()
Fny = TblFny(T, D)
If AyIsEmpty(Fny) Then Exit Function
Dim SeqNo%
SeqNo = 0
For Each F In Fny
    Erase Dr
    Push Dr, T
    Push Dr, SeqNo: SeqNo = SeqNo + 1
    PushAy Dr, FldInfDr(T, F, D)
    Push O, Dr
Next
'DryAddBrkDr O
TblFldInfDry = O
End Function

Function TblFldInfFny() As String()
Dim O$()
Push O, "Tbl"
Push O, "SeqNo"
PushAy O, FldInfFny
TblFldInfFny = O
End Function

Function TblFlds(T, Optional D As Database) As Dao.Fields
Set TblFlds = Tbl(T, D).Fields
End Function

Function TblFny(T, Optional D As Database) As String()
TblFny = FldsFny(Tbl(T, D).Fields)
End Function

Function TblNxtId&(T, Optional F)
Dim S$: S = FmtQQ("select Max(?) from ?", Dft(F, T), T)
TblNxtId = SqlLng(S) + 1
End Function

Function TblPk(T, Optional D As Database) As String()
Dim I As Dao.Index, O$(), F
On Error GoTo X
If Tbl(T, D).Indexes.Count = 0 Then Exit Function
On Error GoTo 0
For Each I In Tbl(T, D).Indexes
    If I.Primary Then
        For Each F In I.Fields
            Push O, F.Name
        Next
        TblPk = O
        Exit Function
    End If
Next
X:
End Function

Function TblRecCnt&(T, Optional D As Database)
On Error GoTo X
TblRecCnt = SqlLng(FmtQQ("Select COunt(*) from [?]", T), D)
Exit Function
X:
TblRecCnt = -1
End Function

Function TblStruLin$(T, Optional SkipTblNm As Boolean, Optional D As Database)
Dim O$(): O = TblFny(T, D): If AyIsEmpty(O) Then Exit Function
O = FnyQuote(O, TblPk(T, D))
O = FnyQuoteIfNeed(O)
Dim J%, V
J = 0
For Each V In O
    O(J) = Replace(V, T, "*")
    J = J + 1
Next
If SkipTblNm Then
    TblStruLin = JnSpc(O)
Else
    TblStruLin = T & " = " & JnSpc(O)
End If
End Function

Function TmpDb() As Database
Set TmpDb = DBEngine.CreateDatabase(TmpFb, Dao.LanguageConstants.dbLangGeneral)
End Function

Private Sub TblPk__Tst()
Dim Dr(), Dry(), T
For Each T In DbTny
    Erase Dr
    Push Dr, T
    PushAy Dr, TblPk(T)
    Push Dry, Dr
Next
DryBrw Dry
End Sub

Sub Tst()
TblPk__Tst
End Sub
