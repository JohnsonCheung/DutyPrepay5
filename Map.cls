VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Map"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Private Sy1_$()
Private Sy2_$()
Function Init(Optional Sy1, Optional Sy2) As Map
If Not IsMissing(Sy1) Then AsrtIsSy Sy1
If Not IsMissing(Sy2) Then AsrtIsSy Sy2
AsrtIsEq UB(Sy1), UB(Sy2)
Sy1_ = Sy1
Sy2_ = Sy2
Set Init = Me
End Function
Function InitByMapStr(MapStr) As Map

End Function
Property Get Sy1() As String()
Sy1 = Sy1_
End Property
Property Get Sy2() As String()
Sy2 = Sy2_
End Property
