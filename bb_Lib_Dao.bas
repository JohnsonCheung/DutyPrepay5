Attribute VB_Name = "bb_Lib_Dao"
Option Compare Database
Option Explicit

Function TmpDb() As Db
Set TmpDb = Db(DBEngine.CreateDatabase(TmpFb, Dao.LanguageConstants.dbLangGeneral))
End Function


Sub AssertT(T, Optional D As Dao.Database)
On Error GoTo X:
Dim A$
A = D.TableDefs(T).Name
Exit Sub
X:
Err.Raise 1, , "Tbl[" & T & "] not found in Db[" & D.Name & "]"
End Sub

Function CurDb() As Db
Static X As Database
If IsNothing(X) Then Set X = CurrentDb
Set CurDb = Db(X)
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

Function DftDb(D As Database) As Database
If IsNothing(D) Then
    Set DftDb = CurDb.Db
Else
    Set DftDb = D
End If
End Function



Function IsNeedQuote(S) As Boolean
IsNeedQuote = True
If HasSubStr(S, " ") Then Exit Function
If HasSubStr(S, "#") Then Exit Function
If HasSubStr(S, ".") Then Exit Function
IsNeedQuote = False
End Function

