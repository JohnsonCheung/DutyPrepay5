Attribute VB_Name = "Ide_CmdBar"
Option Explicit
Option Compare Database

Function CurVbe() As Vbe
Set CurVbe = Application.Vbe
End Function

Function DftVbe(A As Vbe) As Vbe
If IsNothing(A) Then
    Set DftVbe = CurVbe
Else
    Set DftVbe = A
End If
End Function

Function VbeCmdBarAy(Optional A As Vbe) As Office.CommandBar()
Dim O() As Office.CommandBar
Dim I
For Each I In DftVbe(A).CommandBars
    PushObj O, I
Next
VbeCmdBarAy = O
End Function

Function VbeCmdBarNy(Optional A As Vbe) As String()
VbeCmdBarNy = ObjAyPrpSy(VbeCmdBarAy(A), "Name")
End Function
