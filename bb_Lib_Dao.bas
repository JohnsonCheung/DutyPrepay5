Attribute VB_Name = "bb_Lib_Dao"
Option Compare Database
Option Explicit

Sub AddFld(T, F, Ty As DAO.DataTypeEnum, Optional D As DAO.Database)
Dim mFld As New DAO.Field
mFld.Name = F
mFld.Type = Ty
Flds(T, D).Append mFld
End Sub

Sub AssertT(T, Optional D As DAO.Database)
On Error GoTo X:
Dim A$
A = D.TableDefs(T).Name
Exit Sub
X:
Err.Raise 1, , "Tbl[" & T & "] not found in Db[" & D.Name & "]"
End Sub

Sub BrwSql(Sql$, Optional D As DAO.Database)
DrsBrw SqlDrs(Sql, D)
End Sub

Sub BrwTbl(T, Optional D As DAO.Database)
DtBrw TblDt(T, D)
End Sub

Function CurDb() As Database
Static X As Database
If IsNothing(X) Then Set X = CurrentDb
Set CurDb = X
End Function

Function DaoTyStr$(T As DAO.DataTypeEnum)
Dim O$
Select Case T
Case DAO.DataTypeEnum.dbBoolean: O = "Boolean"
Case DAO.DataTypeEnum.dbDouble: O = "Double"
Case DAO.DataTypeEnum.dbText: O = "Text"
Case DAO.DataTypeEnum.dbDate: O = "Date"
Case Else: Stop
End Select
DaoTyStr = O
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

Function Fld(T, F, Optional D As Database) As DAO.Field
Set Fld = Tbl(T, D).Fields(F)
End Function

Property Get Flds(T, Optional D As DAO.Database) As DAO.Fields
Set Flds = Tbl(T, D).Fields
End Property

Function FldsDr(Flds As DAO.Fields) As Variant()
Dim O()
ReDim O(Flds.Count - 1)
Dim J%, F As DAO.Field
For Each F In Flds
    O(J) = F.Value
    J = J + 1
Next
FldsDr = O
End Function

Function FldsFny(Flds As DAO.Fields) As String()
Dim O$()
Dim F As DAO.Field
For Each F In Flds
    Push O, F.Name
Next
FldsFny = O
End Function

Function HasFld(T, F, Optional D As Database) As Boolean
AssertT T, D
HasFld = HasFld_Tbl(Tbl(T, D), F)
End Function

Function HasFld_Flds(Flds As DAO.Fields, F) As Boolean
Dim I As DAO.Field
For Each I In Flds
    If I.Name = F Then HasFld_Flds = True: Exit Function
Next
End Function

Function HasFld_Tbl(T As DAO.TableDef, F) As Boolean
HasFld_Tbl = HasFld_Flds(T.Fields, F)
End Function

Function Tbl(T, Optional D As Database) As DAO.TableDef
Set Tbl = DftDb(D).TableDefs(T)
End Function

Function TblFld(T, F, Optional D As Database) As DAO.Field
Set TblFld = Tbl(T, D).Fields(F)
End Function

Function TblFlds(T, Optional D As Database) As DAO.Fields
Set TblFlds = Tbl(T, D).Fields
End Function

Function TblFny(T, Optional D As Database) As String()
TblFny = FldsFny(Tbl(T, D).Fields)
End Function

Function TblNxtId&(T, Optional F)
Dim S$: S = FmtQQ("select Max(?) from ?", Dft(F, T), T)
TblNxtId = SqlLng(S) + 1
End Function

Function TblStruLin$(T, Optional D As Database)
Dim O$(): O = TblFny(T, D)
Dim J%, V
J = 0
For Each V In O
    O(J) = Replace(V, T, "*")
    J = J + 1
Next
TblStruLin = T & " = " & JnSpc(O)
End Function

Function TmpDb() As Database
Set TmpDb = DBEngine.CreateDatabase(TmpFb, DAO.LanguageConstants.dbLangGeneral)
End Function

Function Tny(Optional A As Database) As String()
Tny = SqlSy("Select Name from MSysObjects where Type in (1,6) and Left(Name,4)<>'MSYS'", A)
End Function
