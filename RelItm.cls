VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RelItm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Public Nm As String
Private Chd_() As String
Public Dta As Variant
Property Get Chd() As String()
Chd = Chd_
End Property
Property Let Chd(Chd$())
Chd_ = Chd
End Property
Function Init(Nm_, Optional Chd, Optional Dta_) As RelItm
Nm = Nm_
If Not IsMissing(Chd) Then Chd_ = Chd
If Not IsMissing(Dta) Then Asg Dta_, Dta
End Function
