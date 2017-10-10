Attribute VB_Name = "Ide_CSub"
Option Explicit
Option Compare Database

Sub MdEnsMthCSub(Optional A As CodeModule)
Dim MthNm
For Each MthNm In MdMthNy(, A)
    MthEnsCSub MthNm, A
Next
End Sub

Sub MthEnsCSub(Optional MthNmOpt, Optional A As CodeModule)
Const CSub$ = "MthEnsCSub"
Dim MthNm$
    MthNm = DftMthNm(MthNmOpt)
    
Dim MthLines$
Dim MthLy$()
Dim MthLno&
    MthLno = MdMthLno(MthNm, A)
    MthLy = MdMthLy(MthNm, A)
    MthLines$ = JnCrLf(MthLy)

Dim IsUsingCSub As Boolean '-> NewAt
    IsUsingCSub = False
    If HasSubStrAy(MthLines, Array("Er CSub,", "Debug.Print CSub", ", CSub,")) Then
        IsUsingCSub = True
    End If

Dim OldCSubIdx%
    Dim J%
    OldCSubIdx = -1
    For J = 0 To UB(MthLy)
        If IsPfx(MthLy(J), "Const CSub") Then
            OldCSubIdx = J
        End If
    Next

Dim OldLno&
    OldLno = IIf( _
        OldCSubIdx >= 0, _
        MthLno + OldCSubIdx, _
        0)

Dim OldLin$
    OldLin = ""
    If OldCSubIdx >= 0 Then
        OldLin = MthLy(OldCSubIdx)
    End If

Dim NewLno&
    NewLno = 0
        If IsUsingCSub Then
            Dim Fnd As Boolean
            For J = 0 To UB(MthLy)
                If LasChr(MthLy(J)) <> "_" Then
                    Fnd = True
                    NewLno = MthLno + J + 1
                    Exit For
                End If
            Next
            If Not Fnd Then Er CSub, "{MthLy} has all lines with _ as sfx with is impossible", MthLy
        End If

Dim NewLin$
    NewLin = ""
        If NewLno > 0 Then
            NewLin = FmtQQ("Const CSub$ = ""?""", MthNm)
        End If
             If OldCSubIdx >= 0 Then OldLin = MthLy(OldCSubIdx) Else OldLin = ""
'do Upd OldLin OldLno NewLin NewLno A
If False Then
    Er CSub, _
        "tracing" & _
        "|{NewLin} is the new CSub line, if any" & _
        "|{NewLno} is Lno the new CSub line#, if any" & _
        "|{OldLin} is the old CSub line, if any" & _
        "|{OldLno} is the old CSub line#", NewLin, NewLno, OldLin, OldLno
End If
With DftMd(A)
    If OldLin = NewLin Then
        Debug.Print FmtQQ("?: Md(?) no change Mth(?)", CSub, MdNm(A), MthNm)
    Else
        If OldLno > 0 Then
            .DeleteLines OldLno         '<==
            Debug.Print FmtQQ("?: Md(?) Mth(?) has line[?] deleted", CSub, MdNm(A), MthNm, OldLin)
        End If
        If NewLno > 0 Then
            .InsertLines NewLno, NewLin  '<==
            Debug.Print FmtQQ("?: Md(?) Mth(?) has line[?] inserted", CSub, MdNm(A), MthNm, NewLin)
        End If
    End If
End With
End Sub

Sub MthEnsCSub__Tst()
MthEnsCSub "MthEnsCSub"
End Sub

