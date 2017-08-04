Attribute VB_Name = "bb_Lib_Dao"
Option Compare Database
Option Explicit
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

Function HasFld_Tbl(T As Dao.TableDef, F) As Boolean
HasFld_Tbl = HasFld_Flds(T.Fields, F)
End Function

Function HasFld(T, F, Optional D As Database) As Boolean
AssertT T, D
HasFld = HasFld_Tbl(Tbl(T, D), F)
End Function

Function DftDb(D As Database) As Database
If IsNothing(D) Then
    Set DftDb = CurDb
Else
    Set DftDb = D
End If
End Function

Function TblNxtId&(T, Optional F)
Dim S$: S = FmtQQ("select Max(?) from ?", Dft(F, T), T)
TblNxtId = SqlLng(S) + 1
End Function

Function CurDb() As Database
Static X As Database
If IsNothing(X) Then Set X = CurrentDb
Set CurDb = X
End Function

Function Tbl(T, Optional D As Database) As Dao.TableDef
Set Tbl = DftDb(D).TableDefs(T)
End Function

Function Fld(T, F, Optional D As Database) As Dao.Field
Set Fld = Tbl(T, D).Fields(F)
End Function
Function FldsFny(Flds As Dao.Fields) As String()
Dim O$()
Dim F As Dao.Field
For Each F In Flds
    Push O, F.Name
Next
FldsFny = O
End Function
Function Tny(Optional A As Database) As String()
Dim O$(), T As TableDef
For Each T In DftDb(A).TableDefs
    Push O, T.Name
Next
Tny = O
End Function
Function TmpDb() As Database
Set TmpDb = DBEngine.CreateDatabase(TmpFb, Dao.LanguageConstants.dbLangGeneral)
End Function
Function TblFlds(T, Optional D As Database) As Dao.Fields
Set TblFlds = Tbl(T, D).Fields
End Function
Function TblFld(T, F, Optional D As Database) As Dao.Field
Set TblFld = Tbl(T, D).Fields(F)
End Function
Sub DrpTbl(T, Optional D As Database)
If IsTbl(T, D) Then D.Execute FmtQQ("Drop Table [?]", T)
End Sub
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
Function DaoTyStr$(T As Dao.DataTypeEnum)
Dim O$
Select Case T
Case Dao.DataTypeEnum.dbBoolean: O = "Boolean"
Case Dao.DataTypeEnum.dbDouble: O = "Double"
Case Dao.DataTypeEnum.dbText: O = "Text"
Case Dao.DataTypeEnum.dbDate: O = "Date"
Case Else: Stop
End Select
DaoTyStr = O
End Function
Function TblFny(T, Optional D As Database) As String()
TblFny = FldsFny(Tbl(T, D).Fields)
End Function
Function HasFld_Flds(Flds As Dao.Fields, F) As Boolean
Dim I As Dao.Field
For Each I In Flds
    If I.Name = F Then HasFld_Flds = True: Exit Function
Next
End Function
Sub AddFld(T, F, Ty As Dao.DataTypeEnum, Optional D As Dao.Database)
Dim mFld As New Dao.Field
mFld.Name = F
mFld.Type = Ty
Flds(T, D).Append mFld
End Sub
Property Get Flds(T, Optional D As Dao.Database) As Dao.Fields
Set Flds = Tbl(T, D).Fields
End Property
Sub AssertT(T, Optional D As Dao.Database)
On Error GoTo X:
Dim A$
A = D.TableDefs(T).Name
Exit Sub
X:
Err.Raise 1, , "Tbl[" & T & "] not found in Db[" & D.Name & "]"
End Sub
