VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Drs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Private Dry_ As Dry
Private Fny_$()  ' as $()
Property Get Dry() As Dry
Set Dry = Dry_
End Property
Property Get Fny() As String()
Fny = Fny_
End Property
Function FldIdxAy(FldNmLvs$) As Long()
'Return Field Idx Ay
FldIdxAy = FnyX.IdxAy(SplitLvs(FldNmLvs))
End Function
Sub AsgFldIdx(FldNmLvs$, ParamArray OAp())
Dim I&(): I = FldIdxAy(FldNmLvs)
Dim J%
For J = 0 To UB(I)
    OAp(J) = I(J)
Next
End Sub
Function FldIdx&(FldNm)
'Return Field Idx Ay
FldIdx = Ay(Fny_).Idx(FldNm)
End Function

Function Init(Fny$(), Optional Dry As Dry) As Drs
Fny_ = Fny
Set Dry_ = DftDry(Dry)
If IsNothing(Dry_) Then Stop
Set Init = Me
End Function
Private Sub DrsLyInsBrkLin__Tst()
Dim TblLy$()
Dim Act$()
Dim Exp$()
'TblLy = FtLy(TstResPth & "DrsLyInsBrkLin.txt")
'Act = LyInsBrkLin(TblLy, "Tbl")
'Exp = FtLy(TstResPth & "DrsLyInsBrkLin_Exp.txt")
'AsrtEqAy Exp, Act
End Sub

Sub Tst()
DrsLyInsBrkLin__Tst
End Sub

Function Clone() As Drs
Set Clone = Nw.Drs(Fny_, Dry_)
End Function
Function LyX(Optional MaxColWdt% = 100) As Ay
Dim O As New Ay
Set LyX = O.Init(Ly(MaxColWdt))
End Function

Function Ly(Optional MaxColWdt% = 100) As String()
If IsEmpty Then Exit Function
Dim O As Drs
Set O = AddRowIdxCol
O.Dry.AddDr O.Fny
Dim Ay As Ay: Set Ay = O.Dry.LyX(MaxColWdt)
Dim Lin$: Lin = Ay.Pop
Dim Hdr$: Hdr = Ay.Pop
Dim OO$()
    PushAy OO, Array(Lin, Hdr)
    PushAy OO, Ay.Ay_
    Push OO, Lin
Ly = OO
End Function
Function Brk(BrkColNm$, Optional MaxColWdt% = 100) As BrkLin
Set Brk = BrkLinByDrsLy(Ly(MaxColWdt), BrkColNm)
End Function
Function IsEmpty() As Boolean
IsEmpty = True
If Dry.NRow = 0 Then Exit Function
If Dry.NCol = 0 Then Exit Function
IsEmpty = False
End Function
Function AddRowIdxCol() As Drs
If Dry.NCol = 0 Then Set AddRowIdxCol = Clone: Exit Function
Dim O As Drs: Set O = Clone
    O.AddFld "RowIdx"
    Dim ODry As Dry: Set ODry = O.Dry
    With ODry
        If .IsEmpty Then GoTo Nxt
        Dim J&, Dr(), I
        For Each I In ODry.Dry
            Dr = I
            Dr(0) = J
            .Dr(J) = Dr
            J = J + 1
            Next
    End With
Nxt:
Set AddRowIdxCol = O
End Function
Function FnyX() As Ay
Set FnyX = Nw.Ay(Fny_)
End Function
Function NRow&()
NRow = Dry_.NRow
End Function
Function URow&()
URow = NRow - 1
End Function
Function AddFld(FldNm, Optional At& = 0) As Drs
Fny_ = FnyX.Ins(FldNm, At).Ay_
Dim J&
If IsEmpty Then Exit Function
For J = 0 To URow
    Dry_.Dr(J) = Dry_.DrX(J).Ins(At:=At).Ay_
Next
End Function
Function WsX(Optional WsNm$ = "Sheet1", Optional Vis As Boolean) As Ws
Dim O As Ws: Set O = NewWsX(WsNm, Vis)
PutAt O.A1
Set WsX = O
End Function

Sub PutAt(At As Range, Optional LoNm$)
Dim A As Rg: Set A = Rg(At)
FnyX.HSq.PutAt At
Dry.Sq.PutAt A.RC(2, 1)
A.WsX.CrtLo LoNm
End Sub

Sub Brw(Optional MaxColWdt% = 100)
LyX(MaxColWdt).Brw
End Sub
Function ColIdx&(Nm$)
ColIdx = FnyX.Idx(Nm)
End Function
Function ColIdxAy(Fny$()) As Long()
ColIdxAy = FnyX.IdxAy(Fny)
End Function
Function Col(ColNm$) As Variant()
Col = Dry_.Col(ColIdx(ColNm))
End Function
Sub FldIdxAsg(FldLvs$, ParamArray OIdxAp())
Dim IdxAy&(): IdxAy = FldIdxAy(FldLvs)
Dim J&
For J = 0 To UB(IdxAy)
    OIdxAp(J) = IdxAy(J)
Next
End Sub

Function DrX(RIdx&) As Ay
Set DrX = Ay(Dry_.DrX(RIdx))
End Function
Function Sel(Fny) As Drs
Dim OFny$(): OFny = Ny(Fny).Ny
Dim IdxAy&(): IdxAy = ColIdxAy(OFny)
Dim ODry As New Dry
    Dim J&
    For J = 0 To URow
        ODry.AddDr DrX(J).Sel(IdxAy)
    Next
Set Sel = Nw.Drs(OFny, ODry)
End Function

Function StrCol(ColNm$) As String()
StrCol = AySy(Col(ColNm))
End Function

Private Sub Sel__Tst()
Sel("MdNm FunNm Mdy Ty").Brw
End Sub

