Attribute VB_Name = "Vb-Vbl"
'Vbl is Vbl.  It is a string without VbCr and VbLf.  It uses | as VbCrLf.  It can be converted to Lines.
Option Compare Database
Option Explicit
Type VblIndent__TstDta
Vbl As String
RstVblNSpc As Integer
FstVblNSpc As Integer
Exp As String
End Type

Sub AssertIsStr(V)
If VarType(V) <> vbString Then Stop
End Sub

Sub AssertIsVbl(V)
AssertIsStr V
If HasSubStr(V, vbCr) Then Stop
If HasSubStr(V, vbLf) Then Stop
End Sub

Sub AssertIsVblAy(V)
AssertIsAy V
If AyIsEmpty(V) Then Exit Sub
Dim Vbl
For Each Vbl In V
    AssertIsVbl Vbl
Next
End Sub

Function IsVbl(Lines) As Boolean
IsVbl = HasSubStr(Lines, "|")
End Function

Function VblByLines$(Lines)
If HasSubStr(Lines, "|") Then Stop
VblByLines = Replace(Lines, vbCrLf, "|")
End Function

Function VblFstLin$(Vbl)

End Function

Function VblIndent$(Vbl, FstVblNSpc%, Optional RstVblNSpc%)
If FstVblNSpc < 0 Then PrmEr
If RstVblNSpc < 0 Then
    If -RstVblNSpc > FstVblNSpc Then Stop
End If
Dim S0$
    S0 = Space(FstVblNSpc)
If Not IsVbl(Vbl) Then
    VblIndent = S0 & Vbl
    Exit Function
End If

Dim S$
    S = Space(FstVblNSpc + RstVblNSpc)
Dim O$()
    O = VblLy(Vbl)
    Dim J%
    O(0) = S0 & O(0)
    For J = 1 To UB(O)
        O(J) = S & O(J)
    Next
VblIndent = JnVBar(O)
End Function

Function VblLasLin$(Vbl)
VblLasLin = AyLasEle(SplitVBar(Vbl))
End Function

Function VblLines$(Vbl)
VblLines = Replace(Vbl, "|", vbCrLf)
End Function

Function VblLy(Vbl) As String()
VblLy = SplitVBar(Vbl)
End Function

Function VblWdt%(Vbl)
VblWdt = AyWdt(VblLy(Vbl))
End Function

Function VblWdt_Lines%(Vbl)
If Not IsVbl(Vbl) Then
    VblWdt_Lines = Len(Vbl)
    Exit Function
End If
Dim Ay$(): Ay = SplitVBar(Vbl)
VblWdt_Lines = AyWdt(Ay)
End Function

Function VblWdt_Ly%(VBarLy)
Dim O%, I
For Each I In VBarLy
    O = Max(O, VblWdt_Lines(I))
Next
VblWdt_Ly = O
End Function

Private Sub VblIndent__Tstr(A As VblIndent__TstDta)
Dim Act$
With A
    Act = VblIndent(.Vbl, .FstVblNSpc, .RstVblNSpc)
    Debug.Assert Act = .Exp
End With
End Sub

Private Sub VblIndent__Tst()
Dim A As VblIndent__TstDta
With A
    .Vbl = "aaa|b|c     d"
    .RstVblNSpc = -2
    .FstVblNSpc = 4
    .Exp = "    aaa|  b|  c     d"
End With
VblIndent__Tstr A
End Sub
