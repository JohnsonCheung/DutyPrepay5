Attribute VB_Name = "Ide_Ren"
Option Explicit
Option Compare Database

Sub MdRmvNmPfx(Pfx, Optional a As CodeModule)
Dim Nm$: Nm = MdNm(a): If Not IsPfx(Nm, Pfx) Then Exit Sub
MdRen RmvPfx(MdNm(a), Pfx), a
End Sub

Sub PjRmvMdNmPfx(Pfx, Optional a As CodeModule)
Dim I, Md As CodeModule
For Each I In PjMdAy(a)
    Set Md = I
    MdRmvNmPfx Pfx, Md
Next
End Sub
Sub MdRen(NewNm$, Optional a As CodeModule)
Dim Nm$: Nm = MdNm(a)
If NewNm = Nm Then Debug.Print FmtQQ("MdRen: New and Old are same [?]", Nm): Exit Sub
If MdIsExist(NewNm, MdPj(a)) Then
    Debug.Print FmtQQ("MdRen: NewNm[?] exist, cannot MdRen(?)", NewNm, MdNm(a))
    Exit Sub
End If
MdCmp(a).Name = NewNm
End Sub
Sub PjRenPfx(FmPfx, ToPfx, Optional a As VBProject)
Dim Ny$()
    Ny = PjMdNy(a)
    Ny = AyFilter(Ny, "IsPfx", FmPfx)
Dim Nm
    For Each Nm In Ny
        MdRenPfx FmPfx, ToPfx, Md(CStr(Nm), a)
    Next

End Sub

Sub MdRenPfx(FmPfx, ToPfx, Optional a As CodeModule)
Dim Nm$: Nm = MdNm(a)
If Not IsPfx(Nm, FmPfx) Then Exit Sub
Dim NewNm$: NewNm = ToPfx & RmvPfx(Nm, FmPfx)
If MdIsExist(NewNm, MdPj(a)) Then
    Debug.Print FmtQQ("MdRenPfx: New-Md-Nm-[?] exist, cannot rename from-pfx-[?] to-pfx-[?] for md-[?]", NewNm, FmPfx, ToPfx, Nm)
    Exit Sub
End If
MdRen NewNm, a
Debug.Print FmtQQ("MdRenPfx: Md-[?] renamed to [?]", Nm, NewNm)
End Sub
