VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Public D As Database
Public T$
Function Init(TblNm, Optional Db As Database) As DbT
Set D = DftDb(Db)
T = TblNm
Set Init = Me
End Function

Function Dt() As Dt
Set Dt = ccNew.Dt(Fny, Dry, T)
End Function
Function Dry() As Dry
Set Dry = Rs.Dry
End Function
Function Fny() As String()
Fny = Flds.Fny
End Function
Function Flds() As Flds
Set Flds = ccNew.Flds(DaoFlds)
End Function
Function Tbl() As Tbl
Set Tbl = ccNew.Tbl(DaoTbl)
End Function
Function Rs() As Rs
Set Rs = Tbl.Rs
End Function
Function Des$()
Des = Prps.Des
End Function

Function NFld%()
NFld = Flds.NFld
End Function
Function FldInfDry() As Dry
Dim O As New Dry, F, Dr()
Dim SeqNo%
SeqNo = 0
Dim J%
Dim mPk$(): mPk = Pk.Ny
Dim FF$(): FF = Fny
If Not AyIsEmpty(FF) Then
    For Each F In FF
        Erase Dr
        Push Dr, T
        Push Dr, SeqNo: SeqNo = SeqNo + 1
        PushAy Dr, Fld(F).InfDr(mPk)
        O.Push Dr
    Next
End If
Set FldInfDry = O
End Function
Function DaoFlds() As Dao.Fields
Set DaoFlds = DaoTbl.Fields
End Function
Function Prps() As Prps
Set Prps = ccNew.Prps(DaoTbl.Properties)
End Function

Function FldInfFny() As String()
Dim O$()
Push O, "Tbl"
Push O, "SeqNo"
PushAy O, Fld(0).InfFny
FldInfFny = O
End Function

Function NxtId&(Optional F)
Dim S$: S = FmtQQ("select Max(?) from ?", Dft(F, T), T)
NxtId = Sql(S).Lng + 1
End Function
Function DaoTbl() As Dao.TableDef
Set DaoTbl = D.TableDefs(T)
End Function
Function Pk() As Ny
Dim I As Dao.Index, O$(), F
On Error GoTo X
Dim T As Dao.TableDef: Set T = DaoTbl
If T.Indexes.Count = 0 Then GoTo X
On Error GoTo 0
For Each I In T.Indexes
    If I.Primary Then
        For Each F In I.Fields
            Push O, F.Name
        Next
        Set Pk = Ny(O)
        Exit Function
    End If
Next
X:
Set Pk = New Ny
End Function

Function IsPkFld(F) As Boolean
IsPkFld = AyHas(Pk.Ny, F)
End Function

Private Sub Pk__Tst()
Dim Dry As New Dry
    Dim Dr(), T, DbT As DbT
    For Each T In Db.DbTy
        Set DbT = T
        Erase Dr
        Push Dr, DbT.T
        PushAy Dr, DbT.Pk
        Dry.Push Dr
    Next
Dry.Brw
End Sub

Sub Tst()
Pk__Tst
End Sub


Sub AddFld(F, Ty As Dao.DataTypeEnum)
Dim mFld As New Dao.Field
mFld.Name = F
mFld.Type = Ty
DaoFlds.Append mFld
End Sub
Function FnyNy() As Ny
Set FnyNy = Ny(Fny)
End Function
Function StruLin$(Optional SkipTblNm As Boolean)
Dim O$(): O = FnyNy.QuoteIfNeed: If AyIsEmpty(O) Then Exit Function
Dim J%, V
J = 0
For Each V In O
    O(J) = Replace(V, T, "*")
    J = J + 1
Next
If SkipTblNm Then
    StruLin = JnSpc(O)
Else
    StruLin = T & " = " & JnSpc(O)
End If
End Function

Function RecCnt&()
On Error Resume Next
Dim S$: S = FmtQQ("Select COunt(*) from [?]", T)
RecCnt = Sql(S).Lng
Exit Function
X:
RecCnt = -1
End Function
Private Function SimTyAy(Optional Fny) As eSimTy()
Dim mFny$(): mFny = DftFny(Fny)
Dim O() As eSimTy
    Dim U%
    ReDim O(U)
    Dim J%, F
    J = 0
    For Each F In Fny
        O(J) = SimTy(DaoFld(F).Type)
        J = J + 1
    Next
SimTyAy = O
End Function

Function FnySimTyAy(Fny$()) As eSimTy()
Dim U%
    U = UB(Fny)
Dim O() As Dao.DataTypeEnum
    ReDim O(U)
    Dim J%, Flds As Dao.Fields
    Set Flds = DaoFlds
    For J = 0 To U
        O(J) = SimTy(Flds(Fny(J)).Type)
    Next
FnySimTyAy = O
End Function
Function LnkTblFx$()
LnkTblFx = TakBet(DaoTbl.Connect, "Database=", ";")
End Function

Private Function DftFny(Fny_) As String()
If IsMissing(Fny_) Then
    DftFny = Fny
Else
    DftFny = Fny_
End If
End Function

Function Sql(Sql_$) As Sql
Set Sql = ccNew.Sql(Sql_, D)
End Function
Function DaoFld(F) As Dao.Field
Set DaoFld = DaoTbl.Fields(F)
End Function
Function Fld(F) As Fld
Set Fld = ccNew.Fld(DaoFld(F))
End Function

Function HasFld(F) As Boolean
HasFld = Tbl.HasFld(F)
End Function

