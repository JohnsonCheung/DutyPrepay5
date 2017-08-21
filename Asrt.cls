VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Asrt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

Sub Eq(V1, V2, Optional Nm1$ = "Exp", Optional Nm2$ = "Act")
If VarType(V1) <> VarType(V2) Then
    Debug.Print FmtQQ("VarType-Of: ?[?] <> ?[?]", Nm1, Nm2, TypeName(V1), TypeName(V2))
    Stop
End If
If IsPrim(V1) Then
    If V1 <> V2 Then
        Debug.Print "Values are diff"
        Debug.Print Nm1 & "[" & V1 & "]"
        Debug.Print Nm2 & "[" & V2 & "]"
        Stop
    End If
End If
If IsAy(V1) Then
    
End If
End Sub
