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
Private IsBlank_ As Boolean
Private IsRmkLin_ As Boolean
Private IsTyLin_ As Boolean
Private IsEnmLin_ As Boolean
Private IsDfnLin_ As Boolean
Property Get IsBlank() As Boolean
IsBlank = IsBlank_
End Property
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
Function Parser() As Parser
Set Parser = Nw.Parser(L)
End Function
Function Init(SrcLin) As SrcLin
L = SrcLin
If Trim(SrcLin) = "" Then
    IsBlank_ = True
Else
    Dim M As Parser: Set M = Parser
    Mdy_ = M.Mdy
    Ty_ = M.Ty: Ty_ = RmvPfx(Ty_, "Property ")
    If Ty_ <> "" Then Nm_ = M.Nm
    IsRmkLin_ = FstChr(LTrim(L)) = "'"
    IsDfnLin_ = Ty_ <> ""
    IsEnmLin_ = Ty_ = "Enum"
    IsTyLin_ = Ty_ = "Type"
End If
Set Init = Me
End Function
Function ChgMdy$(NewMdy$)
Dim A$: If NewMdy <> "" Then A = NewMdy & " "
If Not IsDfnLin Then Stop
Dim M As Parser: Set M = Parser
M.Mdy
ChgMdy = A & M.L
End Function
Sub Tst()
Dim Act As SrcLin: Set Act = SrcLin("Private Function AA()")
Debug.Assert Act.Mdy = "Private"
Debug.Assert Act.Ty = "Function"
Debug.Assert Act.Nm = "AA"
End Sub

Function IsDfnLin() As Boolean
IsDfnLin = IsDfnLin_
End Function
Function IsEnmLin() As String
IsEnmLin = IsEnmLin_
End Function
Function IsTyLin() As Boolean
IsTyLin = IsTyLin_
End Function

