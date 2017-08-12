VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Private Sy_$()
Property Get Sy() As String()
Sy = Sy_
End Property
Function Init(Sy$()) As Sy
Sy_ = Sy
Set Init = Me
End Function
Function Ay() As Ay
Set Ay = ccNew.Ay(Sy)
End Function
Function Lines$()
Lines = JnCrLf(Sy)
End Function
Sub Push(S)
bb_Lib_Vb_Ay.Push Sy_, S
End Sub
Sub PushAy(Sy)
bb_Lib_Vb_Ay.PushAy Sy_, Sy
End Sub

