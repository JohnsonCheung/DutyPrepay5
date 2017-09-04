Attribute VB_Name = "Dta_SimTy"
Option Explicit
Option Compare Database

Enum eSimTy
    eTxt
    eNbr
    eDte
    eLgc
    eOth
End Enum

Function SimTy(T As Dao.DataTypeEnum) As eSimTy
Dim O As eSimTy
Select Case T
Case _
    Dao.DataTypeEnum.dbBigInt, _
    Dao.DataTypeEnum.dbByte, _
    Dao.DataTypeEnum.dbCurrency, _
    Dao.DataTypeEnum.dbDecimal, _
    Dao.DataTypeEnum.dbDouble, _
    Dao.DataTypeEnum.dbFloat, _
    Dao.DataTypeEnum.dbInteger, _
    Dao.DataTypeEnum.dbLong, _
    Dao.DataTypeEnum.dbNumeric, _
    Dao.DataTypeEnum.dbSingle
    O = eNbr
Case _
    Dao.DataTypeEnum.dbChar, _
    Dao.DataTypeEnum.dbGUID, _
    Dao.DataTypeEnum.dbMemo, _
    Dao.DataTypeEnum.dbText
    O = eTxt
Case _
    Dao.DataTypeEnum.dbBoolean
    O = eLgc
Case _
    Dao.DataTypeEnum.dbDate, _
    Dao.DataTypeEnum.dbTimeStamp, _
    Dao.DataTypeEnum.dbTime
    O = eDte
Case Else
    O = eOth
End Select
SimTy = O
End Function

Function SimTyQuoteTp$(a As eSimTy)
Dim O$
Select Case a
Case eTxt: O = "'?'"
Case eNbr, eLgc: O = "?"
Case eDte: O = "#?#"
Case Else
    Er "Given {eSimTy} should be [eTxt eNbr eDte eLgc]", a
End Select
SimTyQuoteTp = O
End Function

