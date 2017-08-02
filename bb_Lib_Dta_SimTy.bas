Attribute VB_Name = "bb_Lib_Dta_SimTy"
Option Compare Database
Option Explicit
Enum eSimTy
    eTxt
    eNbr
    eDte
    eLgc
    eOth
End Enum
Function SimTyQuoteTp$(A As eSimTy)
Dim O$
Select Case A
Case eTxt: O = "'?'"
Case eNbr, eLgc: O = "?"
Case eDte: O = "#?#"
Case Else
    Err.Raise 1, , "Given SimTy[" & A & "] should be [eTxt eNbr eDte eLgc]"
End Select
SimTyQuoteTp = O
End Function
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
