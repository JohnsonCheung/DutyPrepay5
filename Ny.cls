VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ny"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Private A$()
Function Init(Ny_or_NmLvs) As Ny
Dim I: I = Ny_or_NmLvs
If IsMissing(I) Then
ElseIf IsStrAy(I) Then
    A = I
ElseIf IsStr(I) Then
    A = SplitLvs(I)
Else
    Err.Raise 1, , "NyCv: Given [Ny] must be StrAy or Str, but now [" & TypeName(Ny) & "]"
End If
Set Init = Me
End Function
Property Get Ny() As String()
Ny = A
End Property
Function N&()
N = Sz(A)
End Function
Function IsEmpty() As Boolean
IsEmpty = N = 0
End Function
Function QuoteIfNeed() As String()
If IsEmpty Then Exit Function
Dim O$(), J%, F
O = A
For Each F In A
    If IsNmNeedQuote(F) Then O(J) = Quote(F, "'")
    J = J + 1
Next
QuoteIfNeed = O
End Function



