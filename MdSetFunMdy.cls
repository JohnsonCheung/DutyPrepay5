VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MdSetFunMdy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Private Md As Md
Function AA()

End Function
Function Init(A As CodeModule) As MdSetFunMdy
Set Md = ccNew.Md(A)
Set Init = Me
End Function
Sub SetAllFunMdy(Optional AsPub As Boolean)

End Sub
Sub SetFunMdy(FunNy, Mdy$)
Dim I, Lno&
For Each I In Ny(FunNy).Ny
    Lno = Md.FunLno(I)
    If Lno > 0 Then
        SetSrcLinMdy Lno, Mdy
    End If
Next
End Sub
Sub SetSrcLinMdy(Lno&, Mdy$)
Dim NewLin$: Md.SrcLin(Lno).ChgMdy (Mdy)
Md.RplLin Lno, NewLin
End Sub
Function ExportDfnNy() As String()
ExpDfnNy = SplitSpc(ExpDfnLin)
End Function
Function ExportDfnLin$()
Dim J&, L$
Const A$ = "'Export:"
For J = 1 To Md.Md.CountOfDeclarationLines
    L = Trim(Md.Md.Lines(J, 1))
    If IsPfx(L, A) Then ExportDfnLin = RmvPfx(L, A): Exit Function
Next
End Function
Private Function NonPubLnoAy() As Long()
Dim O&(), J&
For J = 1 To Md.Md.CountOfLines
    With Md.SrcLin(J)
        If Not .IsDfnLin Then GoTo Nxt
        If .Mdy = "" Then GoTo Nxt
        If .Mdy = "Public" Then GoTo Nxt
        Push O, J
    End With
Nxt:
Next
End Function
