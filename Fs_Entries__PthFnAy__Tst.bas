Attribute VB_Name = "Fs_Entries__PthFnAy__Tst"
Option Compare Database
Option Explicit
Private Type TstDta
    P          As Variant
    Spec       As String
    ShouldThow As Boolean
    Exp()      As String
End Type

Sub aaa()
MsgBox "AA"
End Sub

Private Function Act(A As TstDta) As String()
With A
    Act = Fs_Entries.PthFnAy(.P, .Spec)
End With
End Function

Private Function ActOpt(A As TstDta) As SyOpt
On Error GoTo X
With A
    ActOpt = SomSy(Act(A))
End With
Exit Function
X:
End Function

Private Function TstDta0() As TstDta
With TstDta0
    .P = "Variant"
    .Spec = ""
    .ShouldThow = False
    .Exp = Sy()
End With
End Function

Private Function TstDta1() As TstDta
With TstDta1
    .P = "Variant"
    .Spec = ""
    .ShouldThow = False
    .Exp = AyOfStr()
End With
End Function

Private Function TstDta2() As TstDta
With TstDta2
    .P = "Variant"
    .Spec = ""
    .ShouldThow = False
    .Exp = Sy()
End With
End Function

Private Function TstDta3() As TstDta
With TstDta3
    .P = "Variant"
    .Spec = ""
    .ShouldThow = False
    .Exp = Sy()
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

Private Sub TstDtaPush(O() As TstDta, I As TstDta)
Dim N&: N = TstDtaSz(O)
ReDim Preserve O(N)
O(N) = I
End Sub

Private Function TstDtaSz%(A() As TstDta)
On Error Resume Next
TstDtaSz = UBound(A) + 1
End Function

Private Function TstDtaUB%(A() As TstDta)
TstDtaUB = TstDtaSz(A) - 1
End Function

Private Sub Tstr(A As TstDta)
Dim M As SyOpt
    M = ActOpt(A)
With A
    If .ShouldThow Then
        If M.Som Then Stop
    Else
        If Not M.Som Then Stop
        AssertActEqExp M.Sy, .Exp
    End If
End With
End Sub

Sub PthFnAy__Tst()
Dim J%
Dim Ay() As TstDta: Ay = TstDtaAy
For J = 0 To TstDtaUB(Ay)
    If J = 0 Then
        Tstr Ay(J)
    End If
Next
End Sub
