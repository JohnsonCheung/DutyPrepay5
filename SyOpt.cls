VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SyOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Public Som As Boolean
Private Sy_$()
Property Get Sy() As String()
Sy = Sy_
End Property
Function Init(Optional Sy) As StrOpt
Set Init = Me
If IsMissing(Sy_) Then Exit Function
Sy_ = Sy
Som = True
End Function

