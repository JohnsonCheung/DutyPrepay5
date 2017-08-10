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
    Dim I, NeedUnEsc As Boolean
    O = QQStr
    For Each I In Av
        If InStr(I, "?") > 0 Then
            NeedUnEsc = True
            I = Replace(I, "?", Chr(255))
        End If
        O = Replace(O, "?", I, Count:=1)
    Next
    If NeedUnEsc Then O = Replace(O, Chr(255), "?")
FmtQQAv = O
End Function
