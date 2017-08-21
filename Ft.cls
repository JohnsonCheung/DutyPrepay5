VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ft"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Public Ft$
Function Init(Ft_) As Ft
Ft = Ft_
Set Init = Me
End Function
Function OpnApp%()
Dim O%: O = FreeFile(1)
Open Ft For Append As #O
OpnApp = O
End Function

Function OpnInp()
Dim O%: O = FreeFile(1)
Open Ft For Input As #O
OpnInp = O
End Function
Function OpnOup%()
Dim O%: O = FreeFile(1)
Open Ft For Output As #O
OpnOup = O
End Function

Function Brw() As Ft
Shell "NotePad """ & Ft & """", vbMaximizedFocus
Set Brw = Me
End Function

Function Ly() As String()
Dim F%: F = OpnInp
Dim L$, O$()
While Not EOF(F)
    Line Input #F, L
    Push O, L
Wend
Close #F
Ly = O
End Function
Function LyX() As Ay
Set LyX = Ay(Ly)
End Function
Function WrtAy(Ay_) As Ft
StrWrt Ay(Ay_).Lines, Ft
Set WrtAy = Me
End Function
Function WrtStr(S) As Ft
StrWrt S, Ft
Set WrtStr = Me
End Function

Function Rmv3Dash(Optional OIsRmv As Boolean) As Ft
'Return true if some 3Dash is remove
Dim mLy$(): mLy = Ly
OIsRmv = False
If AyIsEmpty(Ly) Then Set Rmv3Dash = Me: Exit Function
Dim I, L$, O$()
For Each I In mLy
    L = Brk1(I, "---", NoTrim:=True).S1
    Push O, L
    If L <> I Then OIsRmv = True
    
Next
If OIsRmv Then WrtAy O
FtRmv3Dash = Me
End Function
