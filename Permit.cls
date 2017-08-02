VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Permit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Public Fx$
Private The_PermitId&
Private Sub Class_Initialize()
The_PermitId = TblNxtId("Permit")
End Sub
Sub Import()
If IsEr Then
    WrtER
    MovToER
    Log MsgEr
Else
    BrwDrs ImpPermDrs
    Stop
    InsDt The_InsDt_Of_TblPermit
    InsDt The_InsDt_Of_TblPermitD
    MovToOK
    Log MsgOK
End If
End Sub
Private Function IsEr() As Boolean
IsEr = Not IsEmptyAy(ErMsgAy)
End Function
Private Sub MovToER()
MovFfn Fx, FfnER
End Sub
Private Sub MovToOK()
MovFfn Fx, FfnOK
End Sub
Private Function FfnOK$()
End Function
Private Function FfnER$()
End Function
Private Sub MovFfn(FmFfn$, ToFfn$)
End Sub
Private Sub WrtER()
End Sub
Private Sub AddDr(Dt As Dt, DrObj As Object)
Dim O()
    Dim U%:    U = UB(Dt.Fny)
    ReDim O(U)
    Dim J%
    For J = 0 To U
        O(J) = CallByName(DrObj, Dt.Fny(J), VbGet)
        J = J + 1
    Next
Push Dt.Dry, O
End Sub
Private Function The_Permit_NSku%()

End Function
Private Function The_Permit_Tot@()

End Function
Private Function The_Permit_Qty&()

End Function
Private Function The_PermitNo$()

End Function
Private Function The_InsDt_Of_TblPermit() As Dt
Dim O As Dt
    O.Fny = SplitSpc("Permit PermitNo PermitDate PostDate NSku Qty Tot GLAc GLAcName BankCode ByUsr DteCrt DteUpd IsImport")

If Not IsEr Then
    'Permit = * *No *Date PostDate NSku Qty Tot GLAc GLAcName BankCode ByUsr DteCrt DteUpd IsImport
    
    Dim Dr As New DrPermit
        Dim M As PermitDftVal: M = PermitDftVal
        Dim Dte As Date: Dte = Date
        With Dr
            .PermitNo = The_PermitNo
            .PermitDate = Dte
            .PostDate = Dte
            .NSku% = The_Permit_NSku
            .Qty& = The_Permit_Qty
            .Tot@ = The_Permit_Tot
            .GLAc = M.GLAc
            .GLAcName = M.GLAcName
            .BankCode = M.BankCode
            .ByUsr = M.ByUsr
            .DteCrt = Now
            '.DteUpd
            .IsImport = True
        End With
    AddDr O, Dr
End If
The_InsDt_Of_TblPermit = O
End Function
Private Function The_InsDt_Of_TblPermitD() As Dt
Exit Function

'PermitD = * Permit Sku SeqNo Qty BchNo Rate Amt DteCrt DteUpd
Dim O As Dt
    O.Fny = SplitSpc("PermitD ")
If Not IsEr Then
    Dim M As DrImpPerm
    Dim Seq&: Seq = 0
    Dim I
    For Each I In ImpPermDry
        Set M = I
        Dim Dr As New DrPermitD
            Dim Dte As Date: Dte = Date
            With Dr
                Seq = Seq + 10
                .Permit = The_PermitId
                .Sku = M.Sku
                .SeqNo = Seq
                .Qty = M.Qty
                .BchNo = M.BchNo
                .Rate = M.Rate
                .Amt = M.Amt
                .DteCrt = Now
                '.DteUpd As Date
            End With
        AddDr O, Dr
    Next
End If
The_InsDt_Of_TblPermitD = O
End Function
Function ImpPermWs() As Worksheet
Set ImpPermWs = OpnWb(Fx).Sheets(1)
End Function
Private Function ImpPermDrs() As Drs
Const MapStr$ = "SKU|Des:Material Description|Qty:Order Qty.|PermitNo:Permit Number|BchNo:Batch Number|Amt:Duty Paid Amount"
Dim X As Drs, Z As Boolean
If Not Z Then
    Dim Ws As Worksheet: Set Ws = ImpPermWs
    Z = True
    Dim Sq: Sq = WsSq(ImpPermWs)
    ClsWsNoSav Ws
    X = SqRs(Sq, MapStr)
End If
ImpPermDrs = X
End Function

Private Function ImpPermDry() As DrImpPerm()
End Function

Private Function MsgOK$()
MsgOK = "sdff"
End Function

Private Function MsgEr$()
End Function

Private Function ErMsgAy() As String()
End Function

