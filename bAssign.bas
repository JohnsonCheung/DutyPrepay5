Attribute VB_Name = "bAssign"
Option Compare Database
Option Explicit

Sub rr()
'Aim: For each OH of 'new' batch#, try assign to PermitD->BchNo & create record in SkuB
'     - For those OH batch# cannot find a record in PermitD->(Sku+BchNo), try assign these BchNo to PermitD->BchNo.
'     - When assigning batch#
'       - Sum of PermitD->Qty for Sku+BchNo > OH
'       - All PermitD->Rate of same Sku+BchNo are the same.
'       - For PermitD, use those latest PermitDate first.
'       - For OH     , use those smallest Batch# first.
'Ref: SkuB = Sku BchNo | DutyRateB
'     PermitD = Permit Sku | SeqNo Qty BchNo Rate Amt DteCrt DteUpd
'     Permit  = * *No | *Date PostDate NSku Qty Tot GLAc GLAcName BankCode ByUsr DteCrt DteUpd
'     OH      = YY MM DD YpStk Sku BchNo | Bott Val
DoCmd.SetWarnings False
'DoCmd.RunSQL "Update PermitD set BchNo=Null"
rr_1CrtTmp
With CurrentDb.OpenRecordset("Select * from `#Assign_OH` order by Sku,BchNo")      ' = Sku,BchNo,OH
    While Not .EOF
        rr_2Bch .Fields(0).Value, .Fields(1).Value, .Fields(2).Value
        .MoveNext
    Wend
    .Close
End With
BuildSkuB ' Re-Create record in SkuB
End Sub

Sub XX()
'Aim: Update PermitD->BchNo & create record in SkuB
'     - Each Tax-Paid-OH item's batch# will assign to PermitD->BchNo
'     - Sum of PermitD->Qty for Sku+BchNo > OH
'     - All PermitD->Rate of same Sku+BchNo are the same.
'     - Use those with PermitDate is latest first.
'Ref: SkuB = Sku BchNo | DutyRateB
'     PermitD = Permit Sku | SeqNo Qty BchNo Rate Amt DteCrt DteUpd
'     Permit  = * *No | *Date PostDate NSku Qty Tot GLAc GLAcName BankCode ByUsr DteCrt DteUpd
'     OH      = YY MM DD YpStk Sku BchNo | Bott Val
DoCmd.SetWarnings False
xx_1CrtTmp
With CurrentDb.OpenRecordset("#Assign_SKU")
    While Not .EOF
        SysCmd acSysCmdSetStatus, .Fields(0).Value
        xx_2SKU .Fields(0).Value
        .MoveNext
    Wend
    SysCmd acSysCmdClearStatus
    .Close
End With
End Sub

Private Sub rr_1CrtTmp()
'Aim: Create #Assign_OH from table-OH for those Tax-Paid item with SKU+BchNo not found in table PermitD
'Ref  #Assign_OH  = SKU,BchNo,OH
'     PermitD = Permit Sku | SeqNo Qty BchNo Rate Amt DteCrt DteUpd
Dim mYY As Byte, mMM As Byte, mDD As Byte: SetLatestYYMMDD mYY, mMM, mDD

DoCmd.RunSql Fmt_Str("Select Distinct SKU,BchNo,Sum(Bott) as OH into `#Assign_OH` from OH x" & _
" where YY={0} and MM={1} and DD={2}" & _
" and YpStk in (Select YpStk from YpStk where IsTaxPaid)" & _
" and SKU in (Select SKU from SKU_StkHld where TaxRate is not null)" & _
" group by SKU,BchNo" & _
" having Sum(Bott)<>0" & _
" order by BchNo desc", mYY, mMM, mDD)

DoCmd.RunSql "Update`#Assign_OH` x inner join PermitD a on a.Sku=x.Sku and a.BchNo=x.BchNo set x.OH=Null"
DoCmd.RunSql "Delete from `#Assign_OH` where OH is null"
End Sub

Private Sub rr_2Bch(pSku$, pBchNo$, pOH&)
'Aim: Read PermitD/Permit
'     Find one or more record in PermitD covering the quantity
'     Set PermitD->BchNo & Write to SkuB
'Ref:#Assign_OH  = SKU,BchNo,OH
'    PermitD = Permit Sku | SeqNo Qty BchNo Rate Amt DteCrt DteUpd

'                                                                    Read PermitD in PermitDate Desc for those BchNo=''
Dim aPermitD&(), aRate@(), aQty&(): rr_2Bch_1PermitD pSku, aPermitD, aRate, aQty: If Sz(aPermitD) = 0 Then Exit Sub
Dim J%
Dim mAyIdx%()
Dim mAyPermitD$()
Dim mRate@
mAyPermitD = rr_2Bch_3AyPermitD(pOH, aPermitD, aRate, aQty, mRate) ' Find continuous PermitD's of same rate with quantity can cover bOH(J).
'                                                                  ' After found, return mAyPermitD, Rate, oIdx and set aQty to zero
If Sz(mAyPermitD) > 0 Then
    DoCmd.RunSql Fmt_Str("Update PermitD set BchNo='{0}' where PermitD in ({1}) and SKU='{2}'", pBchNo, Join(mAyPermitD, ","), pSku)
    DoCmd.RunSql Fmt_Str("Insert into SkuB (Sku,BchNo,DutyRateB) values ('{0}','{1}',{2})", pSku, pBchNo, mRate)
End If
End Sub

Private Sub rr_2Bch_1PermitD(pSku$, oPermitD&(), oRate@(), oQty&())
'Aim: Obtain the o* from #Assign_PermitD in MinPermitDate Desc
Dim mN%
With CurrentDb.OpenRecordset(Fmt_Str("Select PermitD,Round(x.Rate,2) as Rate,x.Qty from PermitD x inner join Permit a on a.Permit=x.Permit" & _
                                     " where SKU='{0}' and Nz(BchNo,'')='' order by PermitDate Desc, x.Rate Desc, x.Qty Desc", pSku))
    While Not .EOF
        ReDim Preserve oPermitD(mN)
        ReDim Preserve oRate(mN)
        ReDim Preserve oQty(mN)
        oPermitD(mN) = !PermitD
        oRate(mN) = !Rate
        oQty(mN) = !Qty
        mN = mN + 1
        .MoveNext
    Wend
    .Close
End With
End Sub

Private Function rr_2Bch_3AyPermitD(pOH&, pPermitD&(), pRate@(), pQty&(), ByRef oRate@) As String()
'Aim: Find continuous PermitD's of same rate with quantity can cover bOH(J).
'     After found, return mAyPermitD, oRate, oIdx
Dim mAyIdx%(): mAyIdx = rr_2Bch_3AyPermitD_1AyIdx(pOH, pRate, pQty)
Dim mUB%: mUB = Sz(mAyIdx) - 1: If mUB < 0 Then Exit Function
oRate = pRate(mAyIdx(0))
Dim O$()
ReDim O(mUB)
Dim I%
For I = 0 To UBound(mAyIdx)
    O(I) = pPermitD(mAyIdx(I))
Next
rr_2Bch_3AyPermitD = O
End Function

Private Function rr_2Bch_3AyPermitD_1AyIdx(pOH&, pRate@(), pQty&()) As Integer()
'Aim: Find continuous records of same rate with quantity can cover bOH(J).
'     After found, return oAyIdx and set aQty to zero

Dim mIdx%: mIdx = rr_2Bch_3AyPermitD_1AyIdx_1Idx(pOH, pRate, pQty) ' Find oIdx from which onward, the pQty can cover pOH and having same rate.
If mIdx = -1 Then Exit Function
Dim mRate@: mRate = pRate(mIdx)
Dim O%()
Dim J%
Dim mOH&: mOH = pOH
For J = mIdx To UBound(pQty)
    If pRate(J) <> mRate Then rr_2Bch_3AyPermitD_1AyIdx = O: Exit Function
    mOH = mOH - pQty(J)
    If pQty(J) > 0 Then Push O, J
    If mOH <= 0 Then rr_2Bch_3AyPermitD_1AyIdx = O: Exit Function
Next
Stop ' impossible to reach here.
End Function

Private Function rr_2Bch_3AyPermitD_1AyIdx_1Idx%(pOH&, pRate@(), pQty&())
'Aim: ' Find oIdx from which onward, the oQty can cover pOH and having same rate.
Dim O%: O = rr_2Bch_3AyPermitD_1AyIdx_1Idx_1Nxt(pQty, pRate)
While O <> -1
    Dim mRate@: mRate = pRate(O)
     If rr_2Bch_3AyPermitD_1AyIdx_1Idx_2IsOK(pOH, O, pRate, pQty) Then rr_2Bch_3AyPermitD_1AyIdx_1Idx = O: Exit Function
    O = rr_2Bch_3AyPermitD_1AyIdx_1Idx_1Nxt(pQty, pRate, O + 1, mRate)
Wend
rr_2Bch_3AyPermitD_1AyIdx_1Idx = -1
End Function

Private Function rr_2Bch_3AyPermitD_1AyIdx_1Idx_1Nxt%(pQty&(), pRate@(), Optional pBeg% = 0, Optional p_Rate@ = 0)
'Aim: Start from pBeg, find oIdx so that pQty(oIdx)>0 and pRate(oIdx)<>p_Rate
Dim O%
For O = pBeg To UBound(pQty)
    If pQty(O) <> 0 Then
        If p_Rate <> pRate(O) Then rr_2Bch_3AyPermitD_1AyIdx_1Idx_1Nxt = O: Exit Function
    End If
Next
rr_2Bch_3AyPermitD_1AyIdx_1Idx_1Nxt = -1
End Function

Private Function rr_2Bch_3AyPermitD_1AyIdx_1Idx_2IsOK(pOH&, pIdx%, pRate@(), pQty&()) As Boolean
'Aim: Return true if pIdx is correct index:
'     Use pIdx & pRate() to find mRate
'     From pIdx onward, pQty of records of same Rate can cover pOH return true else false
Dim mRate@: mRate = pRate(pIdx)

Dim J%
Dim mOH&: mOH = pOH
For J = pIdx To UBound(pRate)
    If pRate(J) <> mRate Then Exit Function
    If pQty(J) >= mOH Then rr_2Bch_3AyPermitD_1AyIdx_1Idx_2IsOK = True: Exit Function
    mOH = mOH - pQty(J)
Next
End Function

Private Sub xx_1CrtTmp()
'Aim: Create 4 temp table
'     #Assign_OH` latest Tax-Paid-Item OH from table-OH
'     #Assign_SKU unique SKU from #OH
'     #Assign_Lot  Each from PermitD having same Rate and in seq of date
'     #Assign_LotD Each lot is what PermitD
'Ref  #Assign_OH  = SKU,BchNo,OH
'     #Assign_SKU = SKU
'     #Assign_Lot = Lot SKU Rate MinPermitDate | Qty
'     #Assign_LotD= Lot PermitD |
'     PermitD = Permit Sku | SeqNo Qty BchNo Rate Amt DteCrt DteUpd
'     Permit  = * *No | *Date ...
Dim mYY As Byte, mMM As Byte, mDD As Byte: SetLatestYYMMDD mYY, mMM, mDD
DoCmd.RunSql Fmt_Str("Select Distinct SKU,BchNo,Sum(Bott) as OH into `#Assign_OH` from OH x" & _
" where YY={0} and MM={1} and DD={2}" & _
" and YpStk in (Select YpStk from YpStk where IsTaxPaid)" & _
" and SKU in (Select SKU from SKU_StkHld where IfTaxable='Y')" & _
" group by SKU,BchNo order by BchNo desc ", mYY, mMM, mDD)
DoCmd.RunSql Fmt_Str("Select Distinct SKU into `#Assign_SKU` from `#Assign_OH`")

xDlt.Dlt_Tbl "#Assign_Lot"
xDlt.Dlt_Tbl "#Assign_LotD"
CurrentDb.Execute "Create Table `#Assign_Lot` (Lot Integer, Sku Text(15), Rate Currency, MinPermitDate date, Qty Long," & _
" Constraint PrimaryKey Primary Key (Lot), Constraint `#Assign_Lot` Unique (Sku,Rate,MinPermitDate)) "
CurrentDb.Execute "Create Table `#Assign_LotD` (Lot Integer, PermitD Long, Constraint `#Assign_LotD` unique (Lot,PermitD))"

Dim mLot%: mLot = 1
Dim mLasSku$
Dim mLasRate$
Dim mMinPermitDate As Date
Dim mQty&
Dim mRsLot As Recordset:  Set mRsLot = CurrentDb.TableDefs("#Assign_Lot").OpenRecordset
Dim mRsLotD As Recordset: Set mRsLotD = CurrentDb.TableDefs("#Assign_LotD").OpenRecordset
Dim mRs As Recordset:     Set mRs = CurrentDb.OpenRecordset("Select x.PermitD, Sku, Round(x.Rate,2) as Rate, PermitDate, x.Qty from PermitD x inner join Permit a on a.Permit=x.Permit order by Sku,PermitDate,Rate")
With mRs
    mRsLotD.AddNew
    mRsLotD!Lot = mLot
    mRsLotD!PermitD = !PermitD
    mRsLotD.Update
    
    mLasSku = !Sku
    mLasRate = !Rate
    mQty = !Qty
    mMinPermitDate = !PermitDate
    .MoveNext
    While Not .EOF
        If mLasSku <> !Sku Or mLasRate <> !Rate Then
            mRsLot.AddNew
            mRsLot!Lot = mLot
            mRsLot!Sku = mLasSku
            mRsLot!Rate = mLasRate
            mRsLot!MinPermitDate = mMinPermitDate
            mRsLot!Qty = mQty
            mRsLot.Update
            
            mLot = mLot + 1
            
            mLasSku = !Sku
            mLasRate = !Rate
            mQty = !Qty
            mMinPermitDate = !PermitDate
        Else
            mQty = mQty + !Qty
        End If
        
        mRsLotD.AddNew
        mRsLotD!Lot = mLot
        mRsLotD!PermitD = !PermitD
        mRsLotD.Update
        
        .MoveNext
    Wend
    mRsLot.AddNew
    mRsLot!Lot = mLot
    mRsLot!Sku = mLasSku
    mRsLot!Rate = mLasRate
    mRsLot!MinPermitDate = mMinPermitDate
    mRsLot!Qty = mQty
    mRsLot.Update
End With
End Sub

Private Sub xx_2SKU(pSku$)
If pSku = "1034125" Then Stop
'Aim: Read *OH & *Lot
'     For each *OH find one or more *Lot covering the quantity
'     Set PermitD->BchNo & Write to SkuB
'Ref:#Assign_OH  = SKU,BchNo,OH
'    #Assign_Lot = Lot SKU Rate MinPermitDate | Qty
'    #Assign_LotD= Lot PermitD |
'    PermitD = Permit Sku | SeqNo Qty BchNo Rate Amt DteCrt DteUpd

Dim bBchNo$(), bOH&():          xx_2SKU_1OH pSku, bBchNo, bOH
Dim aLot%(), aRate@(), aQty&(): xx_2SKU_2Lot pSku, aLot, aRate, aQty: If Sz(aLot) = 0 Then Exit Sub
Dim cPermitD$
Dim J%
Dim mLotIdx%
For J = 0 To UBound(bOH)
    mLotIdx = xx_2SKU_3LotIdx(bOH(J), aLot, aQty) ' Find those Lot's quantity can cover bOH(J)
    If mLotIdx >= 0 Then
        Dim mPermitD$(): mPermitD = SqlToAys("Select PermitD from `#Assign_LotD` where Lot=" & aLot(mLotIdx))
        DoCmd.RunSql Fmt_Str("Update PermitD set BchNo='{0}' where PermitD in ({1}) and SKU='{2}'", bBchNo(J), Join(mPermitD, ","), pSku)
        DoCmd.RunSql Fmt_Str("Insert into SkuB (Sku,BchNo,DutyRateB) values ('{0}','{1}',{2})", pSku, bBchNo(J), aRate(mLotIdx))
'    Else
'        Stop
    End If
Next
End Sub

Private Sub xx_2SKU_1OH(pSku$, oBchNo$(), oOH&())
Dim mN%
mN = 0
With CurrentDb.OpenRecordset(Fmt_Str("Select SKU,BchNo,OH from `#Assign_OH` where SKU='{0}' order by OH Desc", pSku))
    While Not .EOF
        ReDim Preserve oBchNo(mN)
        ReDim Preserve oOH(mN)
        oBchNo(mN) = !BchNo
        oOH(mN) = !OH
        mN = mN + 1
        .MoveNext
    Wend
    .Close
End With
End Sub

Private Sub xx_2SKU_2Lot(pSku$, oLot%(), oRate@(), oQty&())
'Aim: Obtain the o* from #Assign_Lot in MinPermitDate Desc
'Ref: #Assign_Lot = Lot SKU Rate MinPermitDate | Qty
Dim mN%
With CurrentDb.OpenRecordset(Fmt_Str("Select Lot,Rate,Qty from `#Assign_Lot` where SKU='{0}' order by MinPermitDate Desc, Qty Desc", pSku))
    While Not .EOF
        ReDim Preserve oLot(mN)
        ReDim Preserve oRate(mN)
        ReDim Preserve oQty(mN)
        oLot(mN) = !Lot
        oRate(mN) = !Rate
        oQty(mN) = !Qty
        mN = mN + 1
        .MoveNext
    Wend
    .Close
End With
End Sub

Private Function xx_2SKU_3LotIdx%(pOH&, pLot%(), pQty&())
'Aim: Return IdxOf pLot()/pQty() which pQty can cover pOH
Dim O%()
Dim J%
For J = 0 To UBound(pQty)
    If pQty(J) >= pOH Then xx_2SKU_3LotIdx = J: Exit Function
Next
xx_2SKU_3LotIdx = -1
End Function

