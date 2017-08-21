VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Assert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Sub AyNotEmpty(Ay)

End Sub
Sub IsEmpty(V)

End Sub
Sub Eq(V1, V2)
If V1 <> V2 Then
    Debug.Print "V1=[" & V1 & "]"
    Debug.Print "V2=[" & V2 & "]"
    Stop
End If
End Sub
Sub Fail()
Er "Fail"
End Sub
Private Sub Er(Msg$, ParamArray Ap())
Dim Av(): Av = Ap
Debug.Print Msg
Stop
End Sub
