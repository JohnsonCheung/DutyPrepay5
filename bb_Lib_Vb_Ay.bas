Attribute VB_Name = "bb_Lib_Vb_Ay"
Option Compare Database
Option Explicit
Function IsEmptyAy(Ay) As Boolean
IsEmptyAy = (Sz(Ay) = 0)
End Function
Sub Push(O, P)
Dim N&: N = Sz(O)
ReDim Preserve O(N)
O(N) = P
End Sub
Sub PushObj(O, P)
Dim N&: N = Sz(O)
ReDim Preserve O(N)
Set O(N) = P
End Sub

Function Sz&(Ay)
On Error Resume Next
Sz = UBound(Ay) + 1
End Function
Function UB&(Ay)
UB = Sz(Ay) - 1
End Function
Sub RmvLasNEle(Ay, Optional NEle% = 1)
ReDim Preserve Ay(UB(Ay) - NEle)
End Sub
Function Pop(Ay)
Pop = LasEle(Ay)
RmvLasNEle Ay
End Function
Function LasEle(Ay)
LasEle = Ay(UB(Ay))
End Function
Function DblQuoteAy(Ay) As String()
Dim O$(), U&, J&
U = UB(Ay)
ReDim Preserve O(U)
For J = 0 To U
    O(J) = """" & Ay(J) & """"
Next
DblQuoteAy = O
End Function
Function AddAyPfx(Ay, Pfx) As String()
Dim O$(), U&, J&
U = UB(Ay)
If U = -1 Then Exit Function
ReDim Preserve O(U)
For J = 0 To U
    O(J) = Pfx & Ay(J)
Next
AddAyPfx = O
End Function
Function AyIdx&(Ay, Itm)
Dim J&
For J = 0 To UB(Ay)
    If Ay(J) = Itm Then AyIdx = J: Exit Function
Next
AyIdx = -1
End Function
Function AyStrAy(Ay) As String()
If IsEmptyAy(Ay) Then Exit Function
Dim U&: U = UB(Ay)
Dim O$()
    Dim J&
    ReDim O(U)
    For J = 0 To U
        O(J) = Ay(J)
    Next
AyStrAy = O
End Function
Sub PushAy(OAy, Ay)
If IsEmptyAy(Ay) Then Exit Sub
Dim I
For Each I In Ay
    Push OAy, I
Next
End Sub
Function BrwAy(Ay)
Dim T$
T = TmpFt
WrtAy Ay, T
BrwFt T
End Function
Sub WrtAy(Ay, Ft)
WrtStr JnCrLf(Ay), Ft
End Sub
Function QuoteAy(Ay, QuoteStr$) As String()
If IsEmptyAy(Ay) Then Exit Function
Dim U&: U = UB(Ay)
Dim O$()
    ReDim O(U)
    Dim J&
    Dim Q1$, Q2$
    With BrkQuote(QuoteStr)
        Q1 = .S1
        Q2 = .S2
    End With
    For J = 0 To U
        O(J) = Q1 + Ay(J) + Q2
    Next
QuoteAy = O
End Function
Function AyIdxAy(Ay, SubAy) As Long()
If IsEmptyAy(SubAy) Then Exit Function
Dim O&()
Dim U&: U = UB(SubAy)
ReDim O(U)
Dim J&
For J = 0 To U
    O(J) = AyIdx(Ay, SubAy(J))
Next
AyIdxAy = O
End Function
Sub DmpAy(Ay)
If IsEmptyAy(Ay) Then Exit Sub
Dim I
For Each I In Ay
    Debug.Print I
Next
End Sub
