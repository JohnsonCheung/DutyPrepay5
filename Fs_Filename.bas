Attribute VB_Name = "Fs_Filename"
Option Compare Database
Option Explicit

Function FfnAddFnSfx(Fn, Sfx)
FfnAddFnSfx = FfnRmvExt(Fn) & Sfx & FfnExt(Fn)
End Function

Function FfnExt$(Ffn)
Dim P%: P = InStrRev(Ffn, ".")
If P = 0 Then Exit Function
FfnExt = Mid(Ffn, P)
End Function

Function FfnFn$(Ffn)
Dim P%: P = InStrRev(Ffn, "\")
If P = 0 Then FfnFn = Ffn: Exit Function
FfnFn = Mid(Ffn, P + 1)
End Function

Function FfnFnn$(Ffn)
FfnFnn = FfnRmvExt(FfnFn(Ffn))
End Function

Function FfnPth$(Ffn)
Dim P%: P = InStrRev(Ffn, "\")
If P = 0 Then Exit Function
FfnPth = Left(Ffn, P)
End Function

Function FfnRmvExt(Fn)
Dim P%: P = InStrRev(Fn, ".")
If P = 0 Then FfnRmvExt = Left(Fn, P): Exit Function
FfnRmvExt = Left(Fn, P - 1)
End Function

Function FfnRplExt$(Ffn, NewExt)
FfnRplExt = FfnRmvExt(Ffn) & NewExt
End Function

Function TmpFb$(Optional Fdr$)
TmpFb = TmpPth(Fdr) & TmpFn(".accdb")
End Function

Function TmpFn$(Ext$)
TmpFn = TmpNm & Ext
End Function

Function TmpFt$(Optional Fdr$)
TmpFt = TmpPth(Fdr) & TmpFn(".txt")
End Function

Function TmpFx$(Optional Fdr$)
TmpFx = TmpPth(Fdr) & TmpFn(".xlsx")
End Function

Function TmpNm$()
Static X&
TmpNm = "T" & Format(Now(), "YYYYMMDD_HHMMSS") & "_" & X
X = X + 1
End Function

Function TmpPth$(Optional Fdr$)
Static X$
If X = "" Then X = Fso.GetSpecialFolder(TemporaryFolder) & "\"
If Fdr = "" Then
    TmpPth = X
Else
    Dim O$
    O = X & Fdr & "\"
    PthEns O
    TmpPth = O
End If
End Function
