VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Public Dic As Dictionary
Function Init(Optional Dic_ As Dictionary) As Dic
If IsMissing(Dic_) Then
    Set Dic = New Dictionary
Else
    Set Dic = Dic_
End If
Set Init = Me
End Function
Function InitByKyVy(Optional Ky, Optional Vy) As Dic
AsrtIsAy Ky
AsrtIsAy Vy
AsrtIsEq UB(Ky), UB(Vy)
Dim O As New Dictionary
If Not AyIsEmpty(Ky) Then
    Dim K, J&
    For Each K In Ky
        O.Add K, Vy(J)
        J = J + 1
    Next
End If
Set InitByKyVy = O
End Function
Function Brw()
Drs.Brw
Set Dic = Me
End Function
Function Keys()
Keys = Dic.Keys
End Function
Function IsEmpty() As Boolean
IsEmpty = Dic.Count = 0
End Function
Function Drs() As Drs
Dim Dry As New Dry, K
If Not IsEmpty Then
    For Each K In Keys
        Dry.AddDr Array(K, Dic(K))
    Next
End If
Dim Fny$(): Fny = SplitSpc("Key Val")
Set Drs = Drs(Fny, Dry)
End Function

