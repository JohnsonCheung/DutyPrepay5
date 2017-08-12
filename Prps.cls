VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Prps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Private Prps As Dao.Properties
Function Init(A As Dao.Properties) As Prps
Set Prps = A
Set Init = Me
End Function
Function Des$()
Des = Val("Description")
End Function

Function Val(PrpNm)
On Error Resume Next
Val = Prps(PrpNm).Value
End Function
