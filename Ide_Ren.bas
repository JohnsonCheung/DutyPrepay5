Attribute VB_Name = "Ide_Ren"
Option Explicit
Option Compare Database

Sub MdRen(NewNm$, Optional A As CodeModule)
Const CSub$ = "MdRen"
Dim Nm$: Nm = MdNm(A)
If NewNm = Nm Then Er CSub, "Given {Md} name and {NewNm} is same", Nm, NewNm
If MdIsExist(NewNm, MdPj(A)) Then
    Er CSub, "{NewNm} already exist.  Cannot rename {Md}", NewNm, MdNm(A)
End If
MdCmp(A).Name = NewNm
Debug.Print FmtQQ("MdRen: Md-[?] renamed to [?]", Nm, NewNm)
End Sub

Sub MdRenPfx(FmPfx, ToPfx, Optional A As CodeModule)
Const CSub$ = "MdRenPfx"
Dim Nm$: Nm = MdNm(A)
If Not IsPfx(Nm, FmPfx) Then Er CSub, "Given {Md} does have given {Pfx}", MdNm(A), FmPfx
Dim NewNm$: NewNm = ToPfx & RmvPfx(Nm, FmPfx)
MdRen NewNm, A
End Sub

Sub MdRmvNmPfx(Pfx, Optional A As CodeModule)
Dim Nm$: Nm = MdNm(A): If Not IsPfx(Nm, Pfx) Then Exit Sub
MdRen RmvPfx(MdNm(A), Pfx), A
End Sub

Sub PjRenMdPfx(FmMdPfx, ToMdPfx, Optional A As Vbproject)
Dim Ny$()
    Ny = PjMdNy(, A)
    Ny = AyFilter(Ny, "IsPfx", FmMdPfx)
Dim Nm
    For Each Nm In Ny
        MdRenPfx FmMdPfx, ToMdPfx, Md(CStr(Nm), A)
    Next
End Sub

Sub PjRmvMdNmPfx(Pfx, Optional A As CodeModule)
Dim I, Md As CodeModule
For Each I In PjMdAy(, A)
    Set Md = I
    MdRmvNmPfx Pfx, Md
Next
End Sub
