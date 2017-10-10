Attribute VB_Name = "Ide_TstMth"
Option Compare Database
Option Explicit
Private Const PrvMthLns$ = ""
Private Const DftUTstDta% = 3
Public Const TstMdStdMthVbl_TstDtaPush$ = _
"|Private Sub TstDtaPush(O() As TstDta, I As TstDta)" & _
"|Dim N&:N = TstDtaSz(O)" & _
"|ReDim Preserve O(N)" & _
"|O(N) = I" & _
"|End Sub"

Sub AssertIsTstMd(T As CodeModule)
If Not IsTstMd(T) Then
    Er "AssertIsTstMd", "Given {Md} is not a TstMd.  TstMd name must have 3 segments separated by __ and last segment is [Tst]", MdNm(T)
End If
End Sub

Function IsTstMd(A As CodeModule) As Boolean
Dim Ay$(): Ay = Split(MdNm(A), "__")
If Sz(Ay) <> 3 Then Exit Function
If Ay(2) <> "Tst" Then Exit Function
IsTstMd = True
End Function

Sub MdAssertIsNotTst(A As CodeModule)
If IsPfx(MdNm(A), "__Tst") Then Er "Assert_MdNm", "{Md} should not have Sfx '__Tst'", MdNm(A)
End Sub

Sub MdTstMthSetMdy(Mdy$, Optional A As CodeModule)
Dim MthNm, MthNy$()
    MthNy = MdMthNy("*__Tst", A)
    For Each MthNm In MthNy
        TstMthSetMdy MthNm, Mdy, A
    Next
End Sub

Sub MdTstMthSetPrv(Optional A As CodeModule)
MdTstMthSetMdy "Private", A
End Sub

Sub MdTstMthSetPub(Optional A As CodeModule)
MdTstMthSetMdy "Public", A
End Sub

Sub TstMdEns(Optional T As CodeModule)
Dim TstMd As CodeModule
    Set TstMd = DftMd(T)
AssertIsTstMd TstMd
Dim Ay$()
    Ay = Split(TstMd, "__")
Dim MthNm$
    MthNm = Ay(1)
Dim SrcMdNm$
    SrcMdNm = Ay(0)
Dim SrcMd As CodeModule
    Set SrcMd = Md(SrcMdNm, MdPj(TstMd))
TstMthEns MthNm, SrcMd
End Sub

Function TstMdStdMthLines_TstDtaPush$()
TstMdStdMthLines_TstDtaPush = RplVbar(TstMdStdMthVbl_TstDtaPush)
End Function

Sub TstMthEns(Optional SrcMthNmOpt$, Optional A As CodeModule)
Const CSub$ = "TstMthEns"
Dim SrcMd As CodeModule
    Set SrcMd = DftMd(A)
MdAssertIsNotTst SrcMd
Dim Pj As Vbproject
    Set Pj = MdPj(A)

Dim SrcMthNm$
    SrcMthNm = DftMthNm(SrcMthNmOpt)
    If SrcMthNm = "" Then Er CSub, "Given {SrcMthNmOpt} is blank and {CurMthNm} is also blank"
MthNmAssertIsNotTst SrcMthNm

Dim SrcMdNm$
    SrcMdNm = MdNm(SrcMd)
    
Dim TstMdNm$
    TstMdNm = FmtQQ("?__?__Tst", SrcMdNm, SrcMthNm)
    
Dim TstSrc$()
    If PjHasMd(TstMdNm, Pj) Then
        TstSrc = MdLy(Md(TstMdNm))
    End If

Dim SrcMthLin$
    SrcMthLin = MdMthLin(SrcMthNm, SrcMd)
    AssertIsMthLin SrcMthLin
    
Dim NewTstMdLy$()
    NewTstMdLy = TstSrcUpd(TstSrc, SrcMdNm, SrcMthLin)
'----
Dim OldTstMdLines$
    OldTstMdLines = MdLines(Md)

Dim NewTstMdLines$
    NewTstMdLines = JnCrLf(NewTstMdLy)
'=================
If OldTstMdLines = NewTstMdLines Then
    Debug.Print FmtQQ("?: Md(?) has standard TstSrc", CSub, SrcMdNm)
Else
    Debug.Print "Bef MdEns"
    MdRpl NewTstMdLines, MdEns(TstMdNm, , Pj)  '<==
End If
End Sub

Sub TstMthSetMdy(MthNm, Mdy$, Optional A As CodeModule)
AssertIsMdy Mdy
Dim I&
    I = MdMthLno(MthNm, A)
Dim L$
    L = MdLin(I, A)
Dim Old$
    Old = ParseMdy(L)
If Mdy = Old Then Exit Sub
Dim NewL$
    Dim B$
    If Mdy <> "" Then
        B = Mdy & " "
    Else
        B = Mdy
    End If
    NewL = B & L
With DftMd(A)
    .DeleteLines I, 1
    .InsertLines I, NewL
End With
End Sub

Sub TstMthSetPrv(MthNm, Optional A As CodeModule)
TstMthSetMdy MthNm, "Private", A
End Sub

Sub TstMthSetPub(MthNm, Optional A As CodeModule)
TstMthSetMdy MthNm, "", A
End Sub

Function TstSrcUpd(TstSrc$(), SrcMdNm$, SrcMthLin$) As String()
AssertIsMthLin SrcMthLin
Dim Arg As MthArg:              Arg = Fnd_MthArg(SrcMthLin)
Dim PrmAy() As MthPrm:        PrmAy = Arg.PrmAy
Dim HasRetVal As Boolean: HasRetVal = Arg.HasRetVal
Dim RetTy As PrmTy:           RetTy = Arg.RetTy
Dim IsRetAy As Boolean:     IsRetAy = Arg.RetTy.IsAy
Dim UTstDta%:               UTstDta = Fnd_UTstDta(TstSrc)
Dim TDMNy$():                 TDMNy = Fnd_TDM_Ny(PrmAy, HasRetVal)
Dim TDMAsTyAy$():         TDMAsTyAy = Fnd_TDM_AsTyAy(Arg)
Dim TDMIsAy() As Boolean:   TDMIsAy = Fnd_TDM_IsAy(PrmAy, IsRetAy)
Dim TDMDftValAy$():     TDMDftValAy = Fnd_TDM_DftValAy(TDMAsTyAy, TDMIsAy)
Dim SrcMthNm$:             SrcMthNm = Fnd_SrcMthNm(SrcMthLin)
Dim RetTyShtNm$:         RetTyShtNm = Fnd_RetTyShtNm(RetTy)

Dim O$()
    O = TstSrc
    O = Ens_Ty_TstDta(O, TDMNy, TDMIsAy, TDMAsTyAy)
    O = Ens_Mth_TstDtaSz(O)
    O = Ens_Mth_TstDtaUB(O)
    O = Ens_Mth_TstDtaPush(O)
    O = Ens_Mth_TstDtaAy(O, UTstDta)
    O = Ens_Mth_Tstr(O, RetTyShtNm)
    O = Ens_Mth_MainTstMth(O, SrcMthNm)
    O = Ens_Mth_TstDtaI_All(O, UTstDta, TDMNy, TDMDftValAy)
    O = Ens_Mth_Act(O, Arg, SrcMdNm, SrcMthNm)
    O = Ens_Mth_ActOpt(O, HasRetVal, RetTyShtNm)
    O = SrcSrtLy(O)
    O = SrcEnsOptExplicit(O)
    O = SrcEnsOptCmpDb(O)
TstSrcUpd = O
End Function

Private Function Ens_Mth(T$(), MthNm$, NewMthLy$()) As String()
Ens_Mth = SrcAddMthIfNotExist(T, MthNm, NewMthLy)
End Function

Private Function Ens_Mth_Act(T$(), A As MthArg, SrcMdNm$, SrcMthNm$) As String()
Dim MNm$
Dim NewMthLy$()
    MNm = "Act"
    NewMthLy() = Fnd_MthLy_Act(A, SrcMdNm, SrcMthNm)
Ens_Mth_Act = Ens_Mth(T, MNm, NewMthLy)
End Function

Private Function Ens_Mth_ActOpt(T$(), HasRetVal As Boolean, RetTyShtNm$) As String()
Dim MNm$
Dim NewMthLy$()
    MNm = "ActOpt"
    NewMthLy() = Fnd_MthLy_ActOpt(HasRetVal, RetTyShtNm)
Ens_Mth_ActOpt = Ens_Mth(T, MNm, NewMthLy)
End Function

Private Function Ens_Mth_MainTstMth(T$(), SrcMthNm$) As String()
Dim MNm$
Dim NewMthLy$()
    MNm = SrcMthNm & "__Tst"
    NewMthLy() = Fnd_MthLy_MainTstMth(SrcMthNm)
Ens_Mth_MainTstMth = Ens_Mth(T, MNm, NewMthLy)
End Function

Private Function Ens_Mth_TstDtaAy(T$(), UTstDta%) As String()
Const MNm$ = "TstDtaAy"
Dim NewMthLy$()
    NewMthLy() = Fnd_MthLy_TstDtaAy(UTstDta)
Ens_Mth_TstDtaAy = Ens_Mth(T, MNm, NewMthLy)
End Function

Private Function Ens_Mth_TstDtaI(T$(), I%, TstDtaMbrNy$(), TstDtaMbrDftValAy$()) As String()
Dim MNm$
Dim NewMthLy$()
    MNm = "TstDta" & I
    NewMthLy() = Fnd_MthLy_TstDtaI(I, TstDtaMbrNy, TstDtaMbrDftValAy)
Ens_Mth_TstDtaI = Ens_Mth(T, MNm, NewMthLy)
End Function

Private Function Ens_Mth_TstDtaI_All(T$(), UTstDta%, TstDtaMbrNy$(), TstDtaMbrDftValAy$()) As String()
Dim O$()
    O = T
    Dim J%
    For J = 0 To UTstDta
        O = Ens_Mth_TstDtaI(O, J, TstDtaMbrNy, TstDtaMbrDftValAy)
    Next
Ens_Mth_TstDtaI_All = O
End Function

Private Function Ens_Mth_TstDtaPush(T$()) As String()
Const MNm$ = "TstDtaPush"
Dim NewMthLy$()
    NewMthLy() = Fnd_MthLy_TstDtaPush
Ens_Mth_TstDtaPush = Ens_Mth(T, MNm, NewMthLy)
End Function

Private Function Ens_Mth_TstDtaSz(T$()) As String()
Const MNm$ = "TstDtaSz"
Dim NewMthLy$()
    NewMthLy() = Fnd_MthLy_TstDtaSz
Ens_Mth_TstDtaSz = Ens_Mth(T, MNm, NewMthLy)
End Function

Private Function Ens_Mth_TstDtaUB(T$()) As String()
Const MNm$ = "TstDtaUB"
Dim NewMthLy$()
    NewMthLy() = Fnd_MthLy_TstDtaUB
Ens_Mth_TstDtaUB = Ens_Mth(T, MNm, NewMthLy)
End Function

Private Function Ens_Mth_Tstr(T$(), RetTyShtNm$) As String()
Const MNm$ = "Tstr"
Dim NewMthLy$()
    NewMthLy() = Fnd_MthLy_Tstr(RetTyShtNm)
Ens_Mth_Tstr = Ens_Mth(T, MNm, NewMthLy)
End Function

Private Function Ens_Ty(T$(), TyNm$, NewTyLy$()) As String()
Ens_Ty = SrcEnsTy(T, TyNm, NewTyLy)
End Function

Private Function Ens_Ty_TstDta(T$(), TstDtaMbrNy$(), TstDtaMbrIsAy() As Boolean, TstDtaMbrAsTyNy$()) As String()
Const TNm$ = "TstDta"
Dim NewMthLy$()
    NewMthLy() = Fnd_TyLy_TstDta(T, TstDtaMbrNy, TstDtaMbrIsAy, TstDtaMbrAsTyNy)
Ens_Ty_TstDta = Ens_Ty(T, TNm, NewMthLy)
End Function

Private Function Fnd_AsTyNm$(A As PrmTy)
Fnd_AsTyNm = PrmTyAsTyNm(A)
End Function

Private Function Fnd_CallingArgStr$(A As MthArg)
Dim Ny$()
    Ny = MthArgPrmNy(A)
    Ny = AyAddPfx(Ny, ".")
Fnd_CallingArgStr = Quote(JnCommaSpc(Ny), "()")
End Function

Private Function Fnd_DftVal_ByAsTyNm$(AsTyNm$, IsRetAy As Boolean)
Dim O$
If IsRetAy Then
    Select Case AsTyNm
    Case "Boolean": O = "False"
    Case "Integer": O = "AyOfInt()"
    Case "Double": O = "AyOfDbl()"
    Case "Long": O = "AyOfLng()"
    Case "LongLong": O = "AyOfLngLng()"
    Case "Single": O = "AyOfSng()"
    Case "Currency": O = "AyOfCur()"
    Case "String": O = "AyOfStr()"
    Case "Date": O = "AyOfDta()"
    Case "Variant": O = "Array()"
    Case Else: O = FmtQQ("""?()""", AsTyNm)
    End Select
Else
    Select Case AsTyNm
    Case "Boolean": O = "False"
    Case "Integer", "Double", "Long", "LongLong", "Single", "Currency": O = "0"
    Case "String": O = """"""
    Case "Date": O = "#2017/1/1#"
    Case Else: O = FmtQQ("""?""", AsTyNm)
    End Select
End If
Fnd_DftVal_ByAsTyNm = O
End Function

Private Function Fnd_MthArg(MthLin$) As MthArg
Fnd_MthArg = MthArg(MthLin)
End Function

Private Function Fnd_MthLy_Act(A As MthArg, SrcMdNm$, SrcMthNm$) As String()
Dim O$()
Dim MthTy$, RetTyChr$, RetTyAsNm$, CallingArgStr
    With A
       MthTy = IIf(.HasRetVal, "Function", "Sub")
        RetTyChr = .RetTy.TyChr
        RetTyAsNm = .RetTy.TyAsNm
        If A.RetTy.IsAy Then RetTyAsNm = RetTyAsNm & "()"
        RetTyAsNm = AddPfxIf(RetTyAsNm, " As ")
        CallingArgStr = Fnd_CallingArgStr(A)
    End With
Push O, ""
Push O, FmtQQ("Private ? Act?(A As TstDta)?", MthTy, RetTyChr, RetTyAsNm)
Push O, FmtQQ("With A")
Push O, FmtQQ("    Act = ?.??", SrcMdNm, SrcMthNm, CallingArgStr)
Push O, FmtQQ("End With")
Push O, FmtQQ("End ?", MthTy)
Fnd_MthLy_Act = O
End Function

Private Function Fnd_MthLy_ActOpt(HasRetVal As Boolean, RetTyShtNm$) As String()
Dim O$()
Push O, ""
If HasRetVal Then
    Push O, FmtQQ("Private Function ActOpt(A As TstDta) As ?Opt", RetTyShtNm)
    Push O, FmtQQ("On Error Goto X")
    Push O, FmtQQ("With A")
    Push O, FmtQQ("    ActOpt = Som?(Act(A))", RetTyShtNm)
    Push O, FmtQQ("End With")
    Push O, FmtQQ("Exit Function")
    Push O, FmtQQ("X:")
    Push O, FmtQQ("End Function")
Else
    Push O, FmtQQ("Function ActOpt(A As TstDta) As BoolOpt")
    Push O, FmtQQ("On Error Goto X")
    Push O, FmtQQ("Act A ")
    Push O, FmtQQ("ActOpt = SomBool(True)")
    Push O, FmtQQ("Exit Sub")
    Push O, FmtQQ("X:")
    Push O, FmtQQ("End Function")
End If
Fnd_MthLy_ActOpt = O
End Function

Private Function Fnd_MthLy_MainTstMth(SrcMthNm$) As String()
Dim O$()
Push O, ""
Push O, FmtQQ("Sub ?__Tst()", SrcMthNm)
Push O, FmtQQ("Dim J%")
Push O, FmtQQ("Dim Ay() As TstDta:Ay = TstDtaAy")
Push O, FmtQQ("For J=0 to TstDtaUB(Ay)")
Push O, FmtQQ("    If J = J Then")
Push O, FmtQQ("        Tstr Ay(J)")
Push O, FmtQQ("    End If")
Push O, FmtQQ("Next")
Push O, FmtQQ("End Sub")
Fnd_MthLy_MainTstMth = O
End Function

Private Function Fnd_MthLy_TstDtaAy(UTstDta%) As String()
Dim O$()
    Push O, ""
    Push O, FmtQQ("Private Function TstDtaAy() As TstDta()")
    Push O, FmtQQ("Dim O() As TstDta")
    Dim J%
    For J = 0 To UTstDta
    Push O, FmtQQ("TstDtaPush O,TstDta?", J)
    Next
    Push O, FmtQQ("TstDtaAy = O")
    Push O, FmtQQ("End Function")
Fnd_MthLy_TstDtaAy = O
End Function

Private Function Fnd_MthLy_TstDtaI(I%, TstDtaMbrNy$(), TstDtaMbrDftValAy$()) As String()
Dim O$()
Push O, ""
Push O, FmtQQ("Private Function TstDta?() As TstDta", I)
Push O, FmtQQ("With TstDta?", I)
Dim J%
For J = 0 To UB(TstDtaMbrNy)
Push O, "    ." & TstDtaMbrNy(J) & " = " & TstDtaMbrDftValAy(J)
'If TstDtaMbrNy(J) = "ShouldThow" Then Stop
Next
Push O, "End With"
Push O, "End Function"
Fnd_MthLy_TstDtaI = O
End Function

Private Function Fnd_MthLy_TstDtaPush() As String()
Fnd_MthLy_TstDtaPush = SplitVBar(TstMdStdMthVbl_TstDtaPush)
End Function

Private Function Fnd_MthLy_TstDtaSz() As String()
Const Lines$ = _
"|Private Function TstDtaSz%(A() As TstDta)" & _
"|On Error Resume Next" & _
"|TstDtaSz = UBound(A) + 1" & _
"|End Function"
Fnd_MthLy_TstDtaSz = SplitVBar(Lines)
End Function

Private Function Fnd_MthLy_TstDtaUB() As String()
Const Lines$ = _
"|Private Function TstDtaUB%(A() As TstDta)" & _
"|TstDtaUB = TstDtaSz(A) - 1" & _
"|End Function"
Fnd_MthLy_TstDtaUB = SplitVBar(Lines)
End Function

Private Function Fnd_MthLy_Tstr(RetTyShtNm$) As String()
Dim O$()
Push O, ""
Push O, FmtQQ("Private Sub Tstr(A As TstDta)")
Push O, FmtQQ("Dim M As ?Opt", RetTyShtNm)
Push O, FmtQQ("    M = ActOpt(A)", RetTyShtNm)
Push O, FmtQQ("With A")
Push O, FmtQQ("    If .ShouldThow Then")
Push O, FmtQQ("        If M.Som Then Stop")
Push O, FmtQQ("    Else")
Push O, FmtQQ("        If Not M.Som Then Stop")
Push O, FmtQQ("        AssertActEqExp M.?, .Exp", RetTyShtNm)
Push O, FmtQQ("    End If")
Push O, FmtQQ("End With")
Push O, FmtQQ("End Sub")
Fnd_MthLy_Tstr = O
'Private Sub Tstr(A As TstDta)
'Dim M As ThowMsgOrSy: M = ActTMOSy(A)
'With A
'    If .ThowMsg = "" Then
'        If Not M.Som Then Stop
'        AssertActEqExp M.Sy, .Exp
'    Else
'        If M.Som Then Stop
'        AssertActEqExp M.ThowMsg, .ThowMsg
'    End If
'End With
'End Sub
End Function

Private Function Fnd_RetTyShtNm$(RetTy As PrmTy)
Dim Ay$
Dim O$
    With RetTy
        If .IsAy Then Ay = "Ay"
        Select Case .TyChr
        Case "!": O = "Sng"
        Case "@": O = "Cur"
        Case "#": O = "Dbl"
        Case "$": O = "Str"
        Case "%": O = "Int"
        Case "^": O = "LngLng"
        Case "&": O = "Lng"
        End Select
        If O = "" Then
            O = .TyAsNm
        End If
        If O = "" Then
            O = "Var"
        End If
    End With
    Select Case O
    Case "String": O = "Str"
    Case "Integer": O = "Int"
    Case "Long": O = "Lng"
    Case "Currency": O = "Cur"
    Case "Single": O = "Sng"
    Case "Double": O = "Dbl"
    Case "LongLong": O = "Lng"
    End Select
    O = O & Ay
    If O = "StrAy" Then O = "Sy"
Fnd_RetTyShtNm = O
End Function

Private Function Fnd_SrcMthNm$(MthLin$)
Fnd_SrcMthNm = SrcLinMthNm(MthLin)
End Function

Private Function Fnd_TDM_AsTyAy(A As MthArg) As String()
Dim J%, O$()
With A
    For J = 0 To MthPrmUB(.PrmAy)
        Push O, Fnd_AsTyNm(.PrmAy(J).Ty)
    Next
    Push O, "Boolean"
    If .HasRetVal Then Push O, Fnd_AsTyNm(.RetTy)
End With
Fnd_TDM_AsTyAy = O
End Function

Private Function Fnd_TDM_DftValAy(AsTyNy$(), IsAy() As Boolean) As String()
Dim O$()
Dim U%: U = UB(AsTyNy)
ReDim O(U)
Dim J%
For J = 0 To U
    O(J) = Fnd_DftVal_ByAsTyNm(AsTyNy(J), IsAy(J))
Next
Fnd_TDM_DftValAy = O
End Function

Private Function Fnd_TDM_IsAy(A() As MthPrm, IsRetAy As Boolean) As Boolean()
Dim O() As Boolean
Dim U%: U = MthPrmUB(A)
Dim J%
For J = 0 To U
    Push O, A(J).Ty.IsAy
Next
Push O, False ' for TstDta.ShouldThow, it is not Ay
Push O, IsRetAy ' for TstDta.Exp, it is not Ay
Fnd_TDM_IsAy = O
End Function

Private Function Fnd_TDM_Ny(A() As MthPrm, HasRetVal As Boolean) As String()
Dim J%, O$()
For J = 0 To MthPrmUB(A)
    Push O, A(J).Nm
Next
Push O, "ShouldThow"
If HasRetVal Then Push O, "Exp"
Fnd_TDM_Ny = O
End Function

Private Function Fnd_TstMd(A As CodeModule)
Dim Nm$
    Nm = MdNm(A) & "__Tst"
Set Fnd_TstMd = Md(Nm, MdPj(A))
End Function

Private Function Fnd_TyLy_TstDta(T$(), TstDtaMbrNy$(), TstDtaMbrIsAy() As Boolean, TstDtaMbrAsTyNy$()) As String()
Dim U%
    U = UB(TstDtaMbrNy)
Dim Ny$()
    Dim J%
    For J = 0 To U
        Push Ny, TstDtaMbrNy(J) & IIf(TstDtaMbrIsAy(J), "()", "")
    Next
    Ny = AyAlignL(Ny)
Dim O$()
Push O, FmtQQ("Private Type TstDta")
For J = 0 To U
Push O, FmtQQ("    ? As ?", Ny(J), TstDtaMbrAsTyNy(J))
Next
Push O, FmtQQ("End Type")
Fnd_TyLy_TstDta = O
End Function

Private Function Fnd_UTstDta%(T$())
Dim Ny$()
    Ny = SrcMthNy(T)
    Ny = AySelPfx(Ny, "TstDta")
    Ny = AyRmvPfx(Ny, "TstDta")
Dim N%()
    Dim I
    If Not AyIsEmpty(Ny) Then
        For Each I In Ny
            PushNoDup N, Val(I)
        Next
    End If
Dim O%
    O = AyMax(N)
    If O = 0 Then O = DftUTstDta
Fnd_UTstDta = O
End Function

Sub TstMthEns__Tst()
TstMthEns "aa", Md("aaa")
End Sub

Sub TstSrcUpd__Tst()
Dim TstSrc$()
Dim SrcMdNm$
Dim SrcMthLin$
    Push TstSrc, "Option Explicit"
    Push TstSrc, "Option Compare Database"
    SrcMdNm = "Ide_TstMthGen"
    SrcMthLin = "Function TstSrcUpd(TstSrc$(), SrcMdNm$, SrcMthLy$()) As String()"
    
Dim Act$()
    Act = TstSrcUpd(TstSrc, SrcMdNm, SrcMthLin)
AyBrw Act
End Sub

