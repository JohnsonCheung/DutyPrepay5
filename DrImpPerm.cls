VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DrImpPerm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Public Sku$
Public Des$
Public Qty&
Public PermitNo
Public BchNo$
Public Amt@
Property Get Rate@()
If Qty = 0 Then Exit Property
On Error Resume Next
Rate = Amt / Qty
End Property
