VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Private Ay
Function Init(Ay_) As Ay
If Not IsArray(Ay_) Then Stop
Ay = Ay_
Set Init = Me
End Function
Sub Brw()
Dim T$: T = TmpFt
Wrt T
FtBrw T
End Sub
Sub Dmp()
AyDmp Ay
End Sub
Sub Wrt(Ft)
AyWrt Ay, Ft
End Sub
Function Sy() As Sy
Dim O$()
If Not IsEmpty Then
    ReDim O(UB)
    Dim J&, I
    For Each I In Ay
        O(J) = I
    Next
End If
Set Sy = ccNew.Sy(O)
End Function
Function IsObjAy() As Boolean
IsObjAy = VarType(Ay) And vbObject
End Function
Function IsStrAy() As Boolean
IsStrAy = VarType(Ay) And vbString
End Function

Function IsEmpty() As Boolean
IsEmpty = Sz = 0
End Function
Function Sz&()
On Error Resume Next
Sz = UBound(Ay) + 1
End Function
Function UB&()
UB = Sz + 1
End Function
