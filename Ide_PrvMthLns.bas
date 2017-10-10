Attribute VB_Name = "Ide_PrvMthLns"
Option Compare Database
Option Explicit

Function PrvMthLnsByConst$(Optional A As CodeModule)
Dim L&: L = PrvMthLnsConstLno(A)
If L = 0 Then Exit Function
PrvMthLnsByConst = DftMd(A).Lines(L, 1)
End Function

Function PrvMthLnsByMth(Optional A As CodeModule)
Dim L$: L = JnSpc(AySrt(MdPrvMthNy(A)))
If L = "" Then Exit Function
PrvMthLnsByMth = FmtQQ("Private Const PrvMthLns$ = ""?""", L)
End Function

Function PrvMthLnsConstLno&(Optional A As CodeModule)
Dim L, J&
For Each L In MdDclLy(A)
    J = J + 1
    If IsPfx(L, "Private Const PrvMthLns$") Then PrvMthLnsConstLno = J: Exit Function
Next
End Function

Sub PrvMthLnsRfhMd(Optional A As CodeModule)
Dim OldLns$, NewLns$
    OldLns = PrvMthLnsByConst(A)
    NewLns = PrvMthLnsByMth(A)
If OldLns = NewLns Then
    Debug.Print FmtQQ("PrvMthLnsRfh: Module(?) <-- Same", MdNm(A))
    Exit Sub
End If
Stop
PrvMthLnsRmv A
MdAppDclLin NewLns, A
End Sub

Sub PrvMthLnsRfhPj(Optional A As Vbproject)
Dim I, Md As CodeModule
For Each I In PjMdAy(, A)
    Set Md = I
    PrvMthLnsRfhMd Md
Next
End Sub

Sub PrvMthLnsRmv(Optional A As Module)
Dim L%: L = PrvMthLnsConstLno(A)
If L = 0 Then
    Debug.Print FmtQQ("PrvMthLnsRmv: Module(?) No Const Lno", MdNm(A))
    Exit Sub
End If
DftMd(A).DeleteLines L, 1
Debug.Print FmtQQ("PrvMthLnsRmv: Module(?) @Lno(?) is removed", MdNm(A), L)
End Sub
