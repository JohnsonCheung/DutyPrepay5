Attribute VB_Name = "SalRpt_T_Div"
Option Compare Database
Option Explicit
Const FldLvs$ = "Div Nm Seq Sts"
Const EDiv$ = "Dept + Division"
Const ENm$ = "DivNm"
Const ESeq$ = "Seq"
Const ESts$ = "Status"
Dim Prm As SR_Prm
Private Sub SR_T_Div_Sql__Tst()
Debug.Print SR_T_Div_Sql
End Sub

Function SR_T_Div_Sql$()
Prm = SR_Prm
SR_T_Div_Sql = FmtQQVBar("Select|?|  Into #Div|  From Division", Fld)
End Function

Private Function ExprVblAy() As String()
ExprVblAy = Sy(EDiv, ENm, ESeq, ESts)
End Function

Private Function Fld$()
Fld = SqpSel(Fny, ExprVblAy)
End Function

Private Function Fny() As String()
Fny = SplitLvs(FldLvs)
End Function
