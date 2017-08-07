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

Function SimTy(T As DAO.DataTypeEnum) As eSimTy
Dim O As eSimTy
Select Case T
Case _
    DAO.DataTypeEnum.dbBigInt, _
    DAO.DataTypeEnum.dbByte, _
    DAO.DataTypeEnum.dbCurrency, _
    DAO.DataTypeEnum.dbDecimal, _
    DAO.DataTypeEnum.dbDouble, _
    DAO.DataTypeEnum.dbFloat, _
    DAO.DataTypeEnum.dbInteger, _
    DAO.DataTypeEnum.dbLong, _
    DAO.DataTypeEnum.dbNumeric, _
    DAO.DataTypeEnum.dbSingle
    O = eNbr
Case _
    DAO.DataTypeEnum.dbChar, _
    DAO.DataTypeEnum.dbGUID, _
    DAO.DataTypeEnum.dbMemo, _
    DAO.DataTypeEnum.dbText
    O = eTxt
Case _
    DAO.DataTypeEnum.dbBoolean
    O = eLgc
Case _
    DAO.DataTypeEnum.dbDate, _
    DAO.DataTypeEnum.dbTimeStamp, _
    DAO.DataTypeEnum.dbTime
    O = eDte
Case Else
    O = eOth
End Select
SimTy = O
End Function

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

