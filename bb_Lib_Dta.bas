Attribute VB_Name = "bb_Lib_Dta"
Option Compare Database
Option Explicit
Function AySel(Ay, IdxAy&())
Dim U&
    U = UB(IdxAy)
Dim O
    O = Ay
    ReDim O(U)
Dim J&
For J = 0 To U
    O(J) = Ay(IdxAy(J))
Next
AySel = O
End Function

Function DtAySz%(DtAy() As Dt)
On Error Resume Next
DtAySz = UBound(DtAy) + 1
End Function
Function FidxAy(Fny$(), FldNmLvs$) As Long()
'Return Field Idx Ay
FidxAy = AyIdxAy(Fny, SplitSpc(FldNmLvs))
End Function

Sub Fiy(Fny$(), FldNmLvs$, ParamArray OAp())
'Fiy=Field Index Array
Dim Ay$(): Ay = SplitSpc(FldNmLvs)
Dim I&(): I = AyIdxAy(Fny, Ay)
Dim J%
For J = 0 To UB(I)
    OAp(J) = I(J)
Next
End Sub


Function IsEmptyDs(A As Ds) As Boolean
IsEmptyDs = IsEmptyDtAy(A.DtAy)
End Function

Function IsEmptyDt(A As Dt) As Boolean
IsEmptyDt = AyIsEmpty(A.Dry)
End Function

Function IsEmptyDtAy(DtAy() As Dt) As Boolean
IsEmptyDtAy = DtAySz(DtAy) = 0
End Function

Function ObjCollSel(ObjColl, PrpNy) As Drs
Dim Ay$()
    Ay = Ny(PrpNy).Ny
Dim Dry As New Dry
    Dim Obj
    If Not IsEmptyColl(ObjColl) Then
        For Each Obj In ObjColl
            Dry.Push ObjSelPrp(Obj, Ay)
        Next
    End If
Set ObjCollSel = ccNew.Drs(Ay, Dry)
End Function

Function ObjSelPrp(Obj, PrpNy$()) As Variant()
Dim U%
    U = UB(PrpNy)
Dim O()
    ReDim O(U)
    Dim J%
    For J = 0 To U
        O(J) = CallByName(Obj, PrpNy(J), VbGet)
    Next
ObjSelPrp = O
End Function


Private Sub ObjCollSel__Tst()
ObjCollSel(DbT("Permit").DaoFlds, "Name Type").Brw
End Sub

Sub Tst()

ObjCollSel__Tst
End Sub
