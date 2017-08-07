Attribute VB_Name = "bb_Lib_Dta_Ins"
Option Compare Database
Option Explicit
Sub InsDs(A As Ds, Optional D As Database)
RunSqlAy InsDsSqlAy(A, D), D
End Sub
Function InsDsSqlAy(A As Ds, Optional D As Database) As String()
If IsEmptyDs(A) Then Exit Function
Dim O$()
Dim J%
For J = 0 To UBound(A.DtAy)
    PushAy O, InsDtSqlAy(A.DtAy(J), D)
Next
InsDsSqlAy = O
End Function
Sub InsDt(A As Dt, Optional D As Database)
RunSqlAy InsDtSqlAy(A, D), D
End Sub
Private Function TblSimTyAy(T, Optional Fny, Optional D As Database) As eSimTy()
Dim mFny$(): mFny = DftFny(T, Fny, D)
Dim O() As eSimTy
    Dim U%
    ReDim O(U)
    Dim J%, F
    J = 0
    For Each F In Fny
        O(J) = SimTy(Fld(T, F, D).Type)
        J = J + 1
    Next
TblSimTyAy = O
End Function

Private Function DftFny(T, Fny, Optional D As Database) As String()
If IsMissing(Fny) Then
    DftFny = TblFny(Fny, D)
Else
    DftFny = Fny
End If
End Function
Private Function InsValTp$(SimTyAy() As eSimTy, Optional D As Database)
Dim U%
    U = UB(SimTyAy)
Dim Ay$()
    ReDim Ay(U)
Dim J%
For J = 0 To U
    Ay(J) = SimTyQuoteTp(SimTyAy(J))
Next
InsValTp = JnComma(Ay)
End Function
Private Function FnySimTyAy(T, Fny$(), Optional D As Database) As eSimTy()
Dim U%
    U = UB(Fny)
Dim O() As DAO.DataTypeEnum
    ReDim O(U)
    Dim J%, Flds As DAO.Fields
    Set Flds = Tbl(T, D).Fields
    For J = 0 To U
        O(J) = SimTy(Flds(Fny(J)).Type)
    Next
FnySimTyAy = O
End Function
Function InsDtSqlAy(Dt As Dt, Optional D As Database) As String()
If IsEmptyDt(Dt) Then Exit Function
Dim SimTyAy() As eSimTy
SimTyAy = FnySimTyAy(Dt.DtNm, Dt.Fny, D)
Dim ValTp$
    ValTp = InsValTp(SimTyAy, D)
Dim Tp$
    Dim T$, F$
    T = Dt.DtNm
    F = JnComma(Dt.Fny)
    Tp = FmtQQ("Insert into [?] (?) values(?)", T, F, ValTp)
Dim O$()
    Dim Dr
    ReDim O(UB(Dt.Dry))
    Dim J%
    J = 0
    For Each Dr In Dt.Dry
        O(J) = FmtQQAv(Tp, Dr)
        J = J + 1
    Next
InsDtSqlAy = O
End Function
Private Sub InsDtSqlAy__Tst()
EnsTbl_Tmp1
Dim Dt As Dt: Dt = TblDt("Tmp1")
Dim O$(): O = InsDtSqlAy(Dt)
Stop
End Sub
Sub CrtTbl(T, FldDclAy, Optional D As Database)
DftDb(D).Execute FmtQQ("Create Table [?] (?)", T, JnComma(FldDclAy))
End Sub
Property Get SampleDt() As Dt
Dim O As Dt
O.DtNm = "Sample"
O.Dry = Array(Array(1))
O.Fny = SplitLvs("A B C")
End Property
