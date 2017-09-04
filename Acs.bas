Attribute VB_Name = "Acs"
Option Explicit
Option Compare Database

Function Acs() As Access.Application
Static X As Access.Application
On Error GoTo XX
Dim A$: A = X.Name
Set Acs = X
Exit Function
XX:
Set X = New Access.Application
Set Acs = X
End Function

Sub DbBrw(D As Database)
Dim N$: N = D.Name
D.Close
FbBrw N
End Sub

Sub FbBrw(Fb$)
Acs.OpenCurrentDatabase Fb
Acs.Visible = True
End Sub
