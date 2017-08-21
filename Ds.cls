VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Private Dic As New Dictionary
Public DsNm$
Function NDt&()
NDt = Dic.Count
End Function
Function UDt&()
UDt = NDt - 1
End Function
Function DtAy() As Dt()
Dim K
If IsEmpty Then Exit Function
Dim O() As Dt, J%
ReDim O(UDt)
For Each K In DtNy
    Set O(J) = Dt(K)
    J = J + 1
Next
DtAy = O
End Function
Function DtNy() As String()
DtNy = AySy(Dic.Keys)
End Function
Function DtNm$(Idx_or_Nm)
Dim I: I = Idx_or_Nm
If IsNbr(I) Then
    DtNm$ = DtNy()(I)
Else
    DtNm = I
End If
End Function
Function Dt(Idx_or_Nm) As Dt
Set Dt = Dic(DtNm(Idx_or_Nm))
End Function
Function IsEmpty() As Boolean
IsEmpty = NDt = 0
End Function
Sub Brw(Optional MaxColWdt% = 100, Optional DsBrkLinMapStr$)
LyX(MaxColWdt, DsBrkLinMapStr).Brw
End Sub
Function LyX(Optional MaxColWdt% = 1000, Optional DsBrkLinMapStr$) As Ay

End Function
Function Ly(Optional MaxColWdt% = 1000, Optional DsBrkLinMapStr$) As String()
Dim O$()
    Push O, "*Ds " & DsNm
Dim Dic As Dictionary ' DicOf_TblNm_to_BrkColNm
    Set Dic = MapDic(BrkMapStr(DsBrkLinMapStr)).Dic
If Not IsEmpty Then
    Dim J%, DtNm$, Dt As Dt, BrkColNm$, I
    For Each I In DtAy
        Set Dt = I
        DtNm$ = Dt.DtNm
        If Dic.Exists(DtNm) Then BrkColNm = Dic(DtNm) Else BrkColNm = ""
        PushAy O, Dt.Ly(MaxColWdt, BrkColNm)
    Next
End If
Ly = O
End Function

Function WbX() As Wb
Dim O As Wb
Set O = NewWbX
With O.FstWs
    .Name = "Ds"
    .Range("A1").Value = DsNm
End With
Dim Dt As Dt, I
If Not IsEmpty Then
    For Each I In DtAy
        Set Dt = I
        Dt.PutWb O
    Next
End If
Set WbX = O
End Function

Sub AddDt(T As Dt)
If HasDt(T.DtNm) Then Err.Raise 1, , FmtQQ("DsAddDt: Ds[?] already has Dt[?]", DsNm, T.DtNm)
Dic.Add T.DtNm, T
End Sub

Function HasDt(DtNm) As Boolean
If IsEmpty Then Exit Function
HasDt = Dic.Exists(DtNm)
End Function
