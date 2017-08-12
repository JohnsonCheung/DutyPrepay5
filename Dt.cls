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
Property Get Fny() As String()
Fny = Fny_
End Property
Function Init(Fny$(), Dry_ As Dry, Optional DtNm_ = "Dt") As Dt
DtNm = DtNm_
Set Dry = Dry_
Fny_ = Fny
Set Init = Me
End Function

Function Ly(Optional MaxColWdt% = 100, Optional BrkColNm$) As Sy
Dim O$()
    Push O, "*Tbl " & DtNm
    If Not IsEmpty Then PushAy O, Drs.Ly(MaxColWdt, BrkColNm).Sy
Set Ly = Sy(O)
End Function
Function IsEmpty() As Boolean
IsEmpty = Dry.IsEmpty
End Function
Function Drs() As Drs
Set Drs = ccNew.Drs(Fny_, Dry)
End Function

Sub Brw(Optional MaxColWdt% = 100)
Ly(MaxColWdt).Ay.Brw
End Sub
Function Xls() As DtXls
Set Xls = DtXls(Me)
End Function

Sub Dmp()
Ly.Ay.Dmp
End Sub

Function CsvLy(A As Dt) As Sy
Dim O$()
Dim QQStr$
Dim Dr
Push O, JnComma(DblAyQuote(A.Fny))
For Each Dr In A.Dry
    Push O, FmtQQAv(QQStr, Dr)
Next
Set CsvLy = Sy(O)
End Function

