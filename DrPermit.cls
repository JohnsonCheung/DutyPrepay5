VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DrPermit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Public Permit%
Public PermitNo%
Public PermitDate As Date
Public PostDate  As Date
Public NSku%
Public Qty&
Public Tot@
Public GLAc$
Public GLAcName$
Public BankCode$
Public ByUsr$
Public DteCrt As Date
Public DteUpd As Date
Public IsImport As Boolean
