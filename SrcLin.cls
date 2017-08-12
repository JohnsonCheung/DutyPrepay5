VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SrcLin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Public L$
Private Ty_$
Private Nm_$
Private Mdy_$
Private IsRmkLin_ As Boolean
Private IsTyLin_ As Boolean
Private IsEnmLin_ As Boolean
Private IsDfnLin_ As Boolean
Property Get IsRmkLin() As Boolean
IsRmkLin = IsRmkLin_
End Property
Property Get Ty$()
Ty = Ty_
End Property
Property Get Nm$()
Nm = Nm_
End Property
Property Get Mdy$()
Mdy = Mdy_
End Property
Function Init(SrcLin) As SrcLin
L = SrcLin
Dim A$: A = SrcLin
Mdy_ = ParseMdy(A)
Ty_ = ParseDfnTy(A): Ty_ = RmvPfx(Ty_, "Property ")
If Ty_ <> "" Then Nm_ = ParseNm(A)
IsRmkLin_ = FstChr(LTrim(L)) = "'"
IsDfnLin_ = Ty_ <> ""
IsEnmLin_ = Ty_ = "Enum"
IsTyLin_ = Ty_ = "Type"
Set Init = Me
End Function
Function ChgMdy$(NewMdy$)
Dim A$: If NewMdy <> "" Then A = NewMdy & " "
If Not IsDfnLin Then Stop
Dim B$: B = L
ParseMdy B
ChgMdy = A & B
End Function
Private Sub Tst()
Dim Act As SrcLin: Act = SrcLin("Private Function AA()")
Debug.Assert Act.Mdy = "Private"
Debug.Assert Act.Ty = "Function"
Debug.Assert Act.Nm = "AA"
End Sub

Property Get IsDfnLin() As Boolean
IsDfnLin = IsDfnLin_
End Property
Property Get IsEnmLin() As String
IsEnmLin = IsEnmLin_
End Property
Property Get IsTyLin() As Boolean
IsTyLin = IsTyLin_
End Property

