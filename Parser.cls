VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Parser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Public L$
Public OriginL$
Function Init(L_) As Parser
L = L_
OriginL = L_
Set Init = Me
End Function

Function Ty$()
Ty = OneOf(ApSy("Function", "Sub", "Property Get", "Property Let", "Property Set", "Type", "Enum"))
End Function

Function Mdy$()
Mdy = OneOf(ApSy("Public", "Private", "Friend"))
End Function
Function Nm$()
Dim J%
J = 1
If Not IsLetter(L) Then GoTo Nxt
For J = 2 To Len(L)
    If Not IsNmChr(Mid(L, J, 1)) Then GoTo Nxt
Next
Nxt:
If J = 1 Then Exit Function
Nm = Left(L, J - 1)
L = Mid(L, J)
End Function
Private Function IsPfx(S, Pfx) As Boolean
IsPfx = (Left(S, Len(Pfx)) = Pfx)
End Function

Function OneOf(OneOfAy$())
Dim I
For Each I In OneOfAy
    If IsPfx(L, CStr(I)) Then L = RmvFstNChr(RmvPfx(L, I)): OneOf = I: Exit Function
Next
End Function


