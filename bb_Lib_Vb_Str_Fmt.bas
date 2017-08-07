Attribute VB_Name = "bb_Lib_Vb_Str_Fmt"
Option Compare Database
Option Explicit
Function FmtQQ$(QQStr$, ParamArray Ap())
Dim Av(): Av = Ap
FmtQQ = FmtQQAv(QQStr, Av)
End Function
Function FmtQQAv$(QQStr$, Av)
If AyIsEmpty(Av) Then FmtQQAv = QQStr: Exit Function
Dim O$
    Dim I
    O = QQStr
    For Each I In Av
        O = Replace(O, "?", I, Count:=1)
    Next
FmtQQAv = O
End Function
