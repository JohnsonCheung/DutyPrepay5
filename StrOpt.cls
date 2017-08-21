VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StrOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Public Som As Boolean
Public Str$
Function Init(Optional Str_) As StrOpt
Set Init = Me
If IsMissing(Str_) Then Exit Function
Str = Str_
Som = True
End Function
