Attribute VB_Name = "SalRpt__SRSqlOMbrWsOpt__Tst"
Option Compare Database
Option Explicit
Private Type TstDta
    Exp As String
    BrkMbr As Boolean
    InclAdr As Boolean
    InclEmail As Boolean
    InclNm As Boolean
    InclPhone As Boolean
End Type
Private Const PrvMthLns$ = "TstDta0 TstDta1 TstDta10 TstDta11 TstDta12 TstDta2 TstDta3 TstDta4 TstDta5 TstDta6 TstDta7 TstDta8 TstDta9 TstDtaAy TstDtaDmp TstDtaDmp0 TstDtaDmp1 TstDtaDmp2 TstDtaDmp3 TstDtaPush"

Private Function TstDta0() As TstDta
With TstDta0
    .BrkMbr = False
    .Exp = ""
End With
End Function

Private Function TstDta1() As TstDta
With TstDta1
    .BrkMbr = True
    .Exp = "Select|    JCMCode                                                        Mbr ,|    DateDiff(Year, Convert(DateTime, JCMDOB, 112), GETDATE())      Age ,|    JCMSex                                                         Sex ,|    JCMStatus                                                      Sts ,|    JCMDist                                                        Dist,|    JCMArea                                                        Area|  Into #MbrDta|  From JCMember|  Where JCMDCode in (Select Mbr From #TxMbr)"
End With
End Function

Private Function TstDta10() As TstDta
With TstDta10
    .BrkMbr = True
    .InclAdr = True
    .Exp = "XX"
End With
End Function

Private Function TstDta11() As TstDta
With TstDta11
    .BrkMbr = True
    .InclAdr = True
    .Exp = "XX"
End With
End Function

Private Function TstDta12() As TstDta
With TstDta12
    .BrkMbr = True
    .InclAdr = True
    .Exp = "XX"
End With
End Function

Private Function TstDta2() As TstDta
With TstDta2
    .BrkMbr = True
    .InclAdr = True
    .Exp = "Select|    JCMCode                                                        Mbr ,|    DateDiff(Year, Convert(DateTime, JCMDOB, 112), GETDATE())      Age ,|    JCMSex                                                         Sex ,|    JCMStatus                                                      Sts ,|    JCMDist                                                        Dist,|    JCMArea                                                        Area,|    Adr-Express-L1|      Adr-Expression-L2                                              Adr |  Into #MbrDta|  From JCMember|  Where JCMDCode in (Select Mbr From #TxMbr)"
End With
End Function

Private Function TstDta3() As TstDta
With TstDta3
    .BrkMbr = True
    .InclAdr = True
    .InclEmail = True
    .Exp = "Select|    JCMCode                                                        Mbr  ,|    DateDiff(Year, Convert(DateTime, JCMDOB, 112), GETDATE())      Age  ,|    JCMSex                                                         Sex  ,|    JCMStatus                                                      Sts  ,|    JCMDist                                                        Dist ,|    JCMArea                                                        Area ,|    Adr-Express-L1|      Adr-Expression-L2                                              Adr  ,|    JCMEmail                                                       Email|  Into #MbrDta|  From JCMember|  Where JCMDCode in (Select Mbr From #TxMbr)"
End With
End Function

Private Function TstDta4() As TstDta
With TstDta4
    .BrkMbr = True
    .InclAdr = True
    .Exp = "XX"
End With
End Function

Private Function TstDta5() As TstDta
With TstDta5
    .BrkMbr = True
    .InclAdr = True
    .Exp = "XX"
End With
End Function

Private Function TstDta6() As TstDta
With TstDta6
    .BrkMbr = True
    .InclAdr = True
    .Exp = "XX"
End With
End Function

Private Function TstDta7() As TstDta
With TstDta7
    .BrkMbr = True
    .InclAdr = True
    .Exp = "XX"
End With
End Function

Private Function TstDta8() As TstDta
With TstDta8
    .BrkMbr = True
    .InclAdr = True
    .Exp = "XX"
End With
End Function

Private Function TstDta9() As TstDta
With TstDta9
    .BrkMbr = True
    .InclAdr = True
    .Exp = "XX"
End With
End Function

Private Function TstDtaAy() As TstDta()
Dim O() As TstDta
TstDtaPush O, TstDta0
TstDtaPush O, TstDta1
TstDtaPush O, TstDta2
TstDtaPush O, TstDta3
TstDtaPush O, TstDta4
TstDtaPush O, TstDta5
TstDtaPush O, TstDta6
TstDtaPush O, TstDta7
TstDtaPush O, TstDta8
TstDtaPush O, TstDta9
TstDtaPush O, TstDta10
TstDtaPush O, TstDta11
TstDtaPush O, TstDta12
TstDtaAy = O
End Function

Private Sub TstDtaDmp(CasNo%)
'Dim D As New Dictionary
'Dim M As TstDta
'    Dim Ay() As TstDta
'    Ay = TstDtaAy
'    M = Ay(CasNo)
'With M
'    D.Add "BrkMbr", .BrkMbr
'    D.Add "*CaseNo", CasNo
'    D.Add "InclAdr", .InclAdr
'    D.Add "InclEmail", .InclEmail
'    D.Add "InclNm", .InclNm
'    D.Add "InclPhone", .InclPhone
'End With
'Dim Exp$
'Dim Act$
'    Exp = M.Exp
'    Act = SR_SqlOMbrWsOpt(
'D.Add "**", IIf(Act = Exp, "Pass", "Fail")
'DicDmp DicSrt(D)
'If Act = Exp Then
'    Debug.Print "Act = Exp ======================================"
'    Debug.Print RplVbar(Act)
'Else
'    Debug.Print "Exp ========================================="
'    Debug.Print RplVbar(Exp)
'    Debug.Print "Act ========================================="
'    Debug.Print RplVbar(Act)
'End If
'AssertActEqExp Act, Exp
End Sub

Private Sub TstDtaDmp0()
TstDtaDmp 0
End Sub

Private Sub TstDtaDmp1()
TstDtaDmp 1
End Sub

Private Sub TstDtaDmp2()
TstDtaDmp 2
End Sub

Private Sub TstDtaDmp3()
TstDtaDmp 3
End Sub

Private Sub TstDtaPush(O() As TstDta, I As TstDta)

End Sub

Private Sub Tstr(A As TstDta)
With A
    AssertActEqExp SR_SqlOMbrWsOpt(.BrkMbr, .InclNm, .InclAdr, .InclEmail, .InclPhone), .Exp
End With
End Sub

Sub SR_O_MbrWsOpt__Tst()
Dim Ay() As TstDta
    Ay = TstDtaAy
Dim J%
For J = 0 To 12
    Tstr Ay(J)
Next
End Sub
