Attribute VB_Name = "Ide_Mdy"
Option Explicit
Option Compare Database

Sub AssertIsMdy(Mdy)
If Not IsMdy(Mdy) Then Er "AssertIsMdy", "{Mdy} is not Mdy.  Valid Mdy is [Private Public Friend *Blank]", Mdy
End Sub

Function IsMdy(Mdy) As Boolean
IsMdy = AyHas(Array("Private", "Public", "Friend", ""), Mdy)
End Function
