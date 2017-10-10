Attribute VB_Name = "SalRpt__SRECrd_SRECrd__Tst"
Option Compare Database
Option Explicit
Private Type TstDta
    CrdTyLvs As String
    CrdPfxTyDry() As Variant
    ShouldThow As Boolean
    Exp As String
End Type

Sub AA2()
SRECrd__Tst
End Sub

Private Function Act(A As TstDta) As String
With A
    Act = SalRpt__SRECrd.SRECrd(.CrdTyLvs, .CrdPfxTyDry)
End With
End Function

Private Function ActOpt(A As TstDta) As StrOpt
On Error GoTo X
ActOpt = SomStr(Act(A))
Exit Function
X:
End Function

Private Function TstDta0() As TstDta
With TstDta0
    .CrdPfxTyDry = ZZCrdPfxTyDry
    .CrdTyLvs = "1 2 3"
    .ShouldThow = False
    .Exp = "Case When|SHMCode Like '134234%' OR|SHMCode Like '12323%'  THEN 1|Else Case When|SHMCode Like '2444%'    OR|SHMCode Like '2443434%' OR|SHMCode Like '24424%'   THEN 2|Else Case When|SHMCode Like '3%' THEN 3|Else 4|End End End "
End With
End Function

Private Function TstDta1() As TstDta
With TstDta1
    .CrdPfxTyDry = ZZCrdPfxTyDry
    .CrdTyLvs = "1"
    .ShouldThow = False
    .Exp = ""
End With
End Function

Private Function TstDta2() As TstDta
With TstDta2
    .CrdPfxTyDry = ZZCrdPfxTyDry
    .CrdTyLvs = ""
    .ShouldThow = True
    .Exp = ""
End With
End Function

Private Function TstDta3() As TstDta
With TstDta3
    .CrdPfxTyDry = ZZCrdPfxTyDry
    .CrdTyLvs = ""
    .ShouldThow = False
    .Exp = ""
End With
End Function

Private Function TstDtaAy() As TstDta()
Dim O() As TstDta
TstDtaPush O, TstDta0
TstDtaPush O, TstDta1
TstDtaPush O, TstDta2
TstDtaPush O, TstDta3
TstDtaAy = O
End Function

Private Sub TstDtaPush(O() As TstDta, M As TstDta)
Dim N&: N = TstDtaSz(O)
ReDim Preserve O(N)
O(N) = M
End Sub

Private Function TstDtaSz&(A() As TstDta)
On Error Resume Next
TstDtaSz = UBound(A) + 1
End Function

Private Sub Tstr(A As TstDta)
Dim M As StrOpt
M = ActOpt(A)
With A
    If .ShouldThow Then
        If M.Som Then Stop
    Else
        If Not M.Som Then Stop
        AssertActEqExp M.Str, .Exp
    End If
End With
End Sub

Private Function ZZCrdPfxTyDry() As Variant()
ZZCrdPfxTyDry = SRCrdPfxTyDry
End Function

Private Function ZZCrdTyLvs$()
ZZCrdTyLvs = "1 2 3"
End Function

Private Sub SRECrd__Tst()
Dim Ay() As TstDta
    Ay = TstDtaAy
Dim J%
For J = 0 To UBound(Ay)
    If J = 0 Then
        Tstr Ay(J)
    End If
Next
End Sub
