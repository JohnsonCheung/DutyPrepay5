Attribute VB_Name = "bb_Lib_Fs_Filename"
Option Compare Database
Option Explicit

Function AddFnSfx(Fn, Sfx)
AddFnSfx = RmvExt(Fn) & Sfx & ".mdb"
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
FfnFnn = RmvExt(FfnFn(Ffn))
End Function

Function FfnPth$(Ffn)
Dim P%: P = InStrRev(Ffn, "\")
If P = 0 Then Exit Function
FfnPth = Left(Ffn, P)
End Function

Function RmvExt(Fn)
Dim P%: P = InStrRev(Fn, ".")
If P = 0 Then RmvExt = Left(Fn, P): Exit Function
RmvExt = Left(Fn, P - 1)
End Function

Function TmpFb$()
TmpFb = TmpPth & TmpFn(".accdb")
End Function

Function TmpFn$(Ext$)
TmpFn = TmpNm & Ext
End Function

Function TmpFt$()
TmpFt = TmpPth & TmpFn(".txt")
End Function

Function TmpNm$()
TmpNm = "T" & Format(Now(), "YYYYMMDD_HHMMSS")
End Function

Function TmpPth$()
Static X$
If X = "" Then X = Fso.GetSpecialFolder(TemporaryFolder) & "\"
TmpPth = X
End Function
