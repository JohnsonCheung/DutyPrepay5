VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BoolOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Public Bool As Boolean
Public Som As Boolean
Function Init(Optional Bool_) As BoolOpt
Set Init = Me
If IsMissing(Bool_) Then Exit Function
Som = True
Bool = Bool_
End Function
