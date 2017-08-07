Attribute VB_Name = "bb_Lib_Vb_Str_FldMapStr"
Option Compare Database
Option Explicit
Type Map
    Sy1() As String
    Sy2() As String
End Type
Function BrkMapStr(MapStr$) As Map
Dim Ay$(): Ay = Split(MapStr, "|")
Dim Ay1$(), Ay2$()
    Dim I
    For Each I In Ay
        With BrkBoth(I, ":")
            Push Ay1, .S1
            Push Ay2, .S2
        End With
    Next
Dim O As Map
    O.Sy1 = Ay1
    O.Sy2 = Ay2
BrkMapStr = O
End Function
Private Sub BrkMapStr__Tst()
End Sub
