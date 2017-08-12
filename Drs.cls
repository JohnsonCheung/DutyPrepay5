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
Private Fny_$()
Property Get Fny() As String()
Fny = Fny_
End Property
Property Get Dry() As Dry
Set Dry = Dry_
End Property
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
TblLy = FtLy(TstResPth & "DrsLyInsBrkLin.txt")
Act = DrsLyInsBrkLin(TblLy, "Tbl")
Exp = FtLy(TstResPth & "DrsLyInsBrkLin_Exp.txt")
AssertEqAy Exp, Act
End Sub

Sub Tst()
DrsLyInsBrkLin__Tst
End Sub

Function Clone() As Drs
Set Clone = ccNew.Drs(Fny_, Dry_)
End Function
Function Ly(Optional MaxColWdt% = 100, Optional BrkColNm$) As Sy
If IsEmpty Then Exit Function
Dim O As Drs
Set O = AddRowIdxCol
O.Dry.Push O.Fny
Dim Ay$(): Ay = O.Dry.Ly(MaxColWdt).Sy
Dim Lin$: Lin = Pop(Ay)
Dim Hdr$: Hdr = Pop(Ay)
Dim OO$()
    PushAy OO, Array(Lin, Hdr)
    PushAy OO, Ay
    Push OO, Lin
If BrkColNm <> "" Then OO = DrsLyInsBrkLin(OO, BrkColNm)
Set Ly = Sy(OO)
End Function
Function Xls() As DrsXls
Set Xls = DrsXls(Me)
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
Sub AddFld(FldNm, Optional At& = 0)
AyIns Fny_, FldNm, At
Dim J&, Dr(), I
If IsEmpty Then Exit Sub
For Each I In Dry_.Dry
    Dr = I
    AyIns Dr, Empty, At
    Dry_.Dr(J) = Dr
    J = J + 1
Next
End Sub
Sub Brw(Optional MaxColWdt% = 100)
Ly(MaxColWdt).Ay.Brw
End Sub
Function ColIdx&(Nm$)
ColIdx = AyIdx(Fny_, Nm)
End Function
Function ColIdxAy(Fny$()) As Long()
ColIdxAy = AyIdxAy(Fny_, Fny)
End Function
Function Col(ColNm$) As Variant()
Col = Dry_.Col(ColIdx(ColNm))
End Function

Function Sel(Fny) As Drs
Dim OFny$(): OFny = Ny(Fny).Ny
Dim IdxAy&(): IdxAy = ColIdxAy(OFny)
Dim ODry As New Dry
    Dim Dr
    For Each Dr In Dry_.Dry
        Dry.Push AySel(Dr, IdxAy)
    Next
Set Sel = ccNew.Drs(OFny, ODry)
End Function

Function StrCol(ColNm$) As String()
StrCol = AySy(Col(ColNm))
End Function

Private Sub Sel__Tst()
Sel("MdNm FunNm Mdy Ty").Brw
End Sub

