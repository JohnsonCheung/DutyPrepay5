Attribute VB_Name = "cc_MdSetFunMdy"
Option Compare Database
Option Explicit
Function MdSetFunMdy(M As CodeModule) As MdSetFunMdy
Dim O As New MdSetFunMdy
Set MdSetFunMdy = O.Init(M)
End Function
