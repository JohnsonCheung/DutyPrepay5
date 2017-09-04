Attribute VB_Name = "Ide_Ren"
Option Explicit
Option Compare Database

Sub MdRen(NewNm$, Optional A As CodeModule)
Dim Nm$: Nm = MdNm(A)
If NewNm = Nm Then Debug.Print FmtQQ("MdRen: New and Old are same [?]", Nm): Exit Sub
If MdIsExist(NewNm, MdPj(A)) Then
    Debug.Print FmtQQ("MdRen: NewNm[?] exist, cannot MdRen(?)", NewNm, MdNm(A))
    Exit Sub
End If
MdCmp(A).Name = NewNm
End Sub

Sub MdRenPfx(FmPfx, ToPfx, Optional A As CodeModule)
Dim Nm$: Nm = MdNm(A)
If Not IsPfx(Nm, FmPfx) Then Exit Sub
Dim NewNm$: NewNm = ToPfx & RmvPfx(Nm, FmPfx)
If MdIsExist(NewNm, MdPj(A)) Then
    Debug.Print FmtQQ("MdRenPfx: New-Md-Nm-[?] exist, cannot rename from-pfx-[?] to-pfx-[?] for md-[?]", NewNm, FmPfx, ToPfx, Nm)
    Exit Sub
End If
MdRen NewNm, A
Debug.Print FmtQQ("MdRenPfx: Md-[?] renamed to [?]", Nm, NewNm)
End Sub

Sub MdRmvNmPfx(Pfx, Optional A As CodeModule)
Dim Nm$: Nm = MdNm(A): If Not IsPfx(Nm, Pfx) Then Exit Sub
MdRen RmvPfx(MdNm(A), Pfx), A
End Sub

Sub PjRenPfx(FmPfx, ToPfx, Optional A As VBProject)
Dim Ny$()
    Ny = PjMdNy(A)
    Ny = AyFilter(Ny, "IsPfx", FmPfx)
Dim Nm
    For Each Nm In Ny
        MdRenPfx FmPfx, ToPfx, Md(CStr(Nm), A)
    Next

End Sub

Sub PjRmvMdNmPfx(Pfx, Optional A As CodeModule)
Dim I, Md As CodeModule
For Each I In PjMdAy(A)
    Set Md = I
    MdRmvNmPfx Pfx, Md
Next
End Sub
