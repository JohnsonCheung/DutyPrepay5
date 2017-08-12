Attribute VB_Name = "bb_Lib_Ide"
Option Compare Database
Option Explicit
Type LnoCnt
    Lno As Long
    Cnt As Long
End Type
Function LnoCnt(Lno&, Cnt&) As LnoCnt
LnoCnt.Lno = Lno
LnoCnt.Cnt = Cnt
End Function
Sub MdCrt(Optional MdNm$, Optional Ty As vbext_ComponentType = vbext_ct_StdModule, Optional A As VBProject)
Dim O As VBComponent: Set O = DftPj(A).VBComponents.Add(Ty)
O.CodeModule.DeleteLines 1, 2
If MdNm <> "" Then O.Name = MdNm
End Sub
Function DftPjNm$(Optional PjNm$)
If PjNm = "" Then
    DftPjNm = DftPj.Name
Else
    DftPjNm = PjNm
End If
End Function
Function CurMd() As Md
Set CurMd = Md(Application.VBE.ActiveCodePane.CodeModule)
End Function

Function CurPj() As Pj
Set CurPj = Pj(Application.VBE.ActiveVBProject)
End Function

Function DftPj(Optional A As VBProject) As VBProject
If IsNothing(A) Then
    Set DftPj = Application.VBE.ActiveVBProject
Else
    Set DftPj = A
End If
End Function
Function DftMdNm$(Nm$)
If Nm = "" Then
    DftMdNm = CurMd.Nm
Else
    DftMdNm = Nm
End If
End Function

Function JnContinueLin(Ly$()) As String()
Dim O$(): O = Ly
Dim J&
For J = UB(O) - 1 To 0 Step -1
    If LasChr(O(J)) = "_" Then
        O(J) = RmvLasNChr(O(J)) & O(J + 1)
        O(J + 1) = ""
    End If
Next
JnContinueLin = O
End Function


Function ParseDfnTy$(OLin$)
ParseDfnTy = ParseOneOf(OLin, ApSy("Function", "Sub", "Property Get", "Property Let", "Property Set", "Type", "Enum"))
End Function

Function ParseMdy$(OLin$)
ParseMdy = ParseOneOf(OLin, ApSy("Public", "Private", "Friend"))
End Function

Function ParseNm$(OLin$)
Dim J%
J = 1
If Not IsLetter(FstChr(OLin)) Then GoTo Nxt
For J = 2 To Len(OLin)
    If Not IsNmChr(Mid(OLin, J, 1)) Then GoTo Nxt
Next
Nxt:
If J = 1 Then Exit Function
ParseNm = Left(OLin, J - 1)
OLin = Mid(OLin, J)
End Function

Function ParseOneOf(OLin$, OneOfAy$())
Dim I
For Each I In OneOfAy
    If IsPfx(OLin, I) Then OLin = RmvFstNChr(RmvPfx(OLin, I)): ParseOneOf = I: Exit Function
Next
End Function
Private Function IsPfx(S, Pfx) As Boolean
IsPfx = (Left(S, Len(Pfx)) = Pfx)
End Function


