VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Public DtNm As String
Private Fny_() As String
Public Dry As Dry
Function URow&()
URow = NRow - 1
End Function
Property Get Fny() As String()
Fny = Fny_
End Property
Function Init(Fny$(), Dry_ As Dry, Optional DtNm_ = "Dt") As Dt
DtNm = DtNm_
Set Dry = Dry_
Fny_ = Fny
Set Init = Me
End Function
Function Brk(BrkColNm$, Optional MaxColWdt% = 100) As BrkLin
Dim O As New BrkLin
Set Brk = O.InitByDrsLy(Ly(MaxColWdt, BrkColNm), BrkColNm)
End Function
Function PutWb(A As Workbook) As Worksheet
PutAt Nw.Wb(A).AddWsX(DtNm).A1
Set PutWb = Nw.Wb(A).Ws(DtNm)
End Function
Function PutAt(At As Range) As Range
Nw.SqByHAy(Fny_).PutAt At
Dry.PutAt At(2, 1)
Set PutAt = At
End Function

Function Ws(Optional Hid As Boolean) As Worksheet
Set Ws = PutAt(NewA1(DtNm, Not Hid)).Parent
End Function

Function WsX(Optional Hid As Boolean) As Ws
Set WsX = Nw.Ws(Ws(Hid))
End Function

Function LyX(Optional MaxColWdt% = 100, Optional BrkColNm$) As Ay
Set LyX = Nw.Ay(Ly(MaxColWdt, BrkColNm))
End Function

Function Ly(Optional MaxColWdt% = 100, Optional BrkColNm$) As String()
Dim O$()
    Push O, "*Tbl " & DtNm
    If Not IsEmpty Then PushAy O, Drs.Ly(MaxColWdt)
Ly = O
End Function

Function IsEmpty() As Boolean
IsEmpty = Dry.IsEmpty
End Function

Function Drs() As Drs
Set Drs = Nw.Drs(Fny_, Dry)
End Function
Function FldLis$()
FldLis = FnyX.JnComma
End Function
Function Brw(Optional MaxColWdt% = 100, Optional BrkColNm$) As Dt
Brk(BrkColNm, MaxColWdt).LyX.Brw
Set Brw = Me
End Function
Function Sample() As Dt
Dim Dry As New Dry: Dry.AddDr Array(1, 2, 3)
Dim Fny$(): Fny = SplitLvs("A B C")
Set Sample = Nw.Dt(Fny, Dry, "Sample")
End Function

Sub Dmp()
LyX.Dmp
End Sub
Function CsvLyX() As Ay
Set CsvLyX = Nw.Ay(CsvLy)
End Function
Function FnyX() As Ay
Set FnyX = Nw.Ay(Fny_)
End Function
Function NRow&()
NRow = Dry.NRow
End Function
Function CsvLy() As String()
Dim O$()
With EmptyStrAy
    .Push FnyX.DblQuoteJnComma
    If IsEmpty Then GoTo X
    Dim QQStr$
        QQStr = Dry.CsvQQStr
    Dim Dr

    For Each Dr In Dry.Dry
        .Push FmtQQAv(QQStr, Dr)
    Next
X:
    O = .Ay_
End With
CsvLy = O
End Function
