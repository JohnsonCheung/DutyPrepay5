Attribute VB_Name = "bb_Lib_Acs"
Option Compare Database
Option Explicit

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
