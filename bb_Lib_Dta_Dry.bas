Attribute VB_Name = "bb_Lib_Dta_Dry"
Option Compare Database
Option Explicit
Function DryLy(Dry) As String()
If IsEmpty(Dry) Then Exit Function
Dim W%(): W = DryWdtAy(Dry)
If IsEmptyAy(W) Then Exit Function
Dim HdrAy$()
    ReDim HdrAy(UB(W))
    Dim J%
    For J = 0 To UB(W)
        HdrAy(J) = String(W(J), "-")
    Next
Dim Hdr$: Hdr = Quote(Join(HdrAy, "-|-"), "|-*-|")
Dim O$()
    Dim Dr
    Push O, Hdr
    For Each Dr In Dry
        Push O, DrLin(Dr, W)
    Next
    Push O, Hdr
DryLy = O
End Function
Function DryCol(Dry, Optional ColIdx% = 0) As Variant()
If IsEmptyAy(Dry) Then Exit Function
Dim O(), Dr
For Each Dr In Dry
    Push O, Dr(ColIdx)
Next
DryCol = O
End Function
Function DryStrCol(Dry, Optional ColIdx% = 0) As String()
DryStrCol = AySy(DryCol(Dry, ColIdx))
End Function
Sub DmpDry(Dry)
DmpAy DryLy(Dry)
End Sub
Private Function DryWdtAy(Dry) As Integer()
If IsEmptyAy(Dry) Then Exit Function
Dim O%()
    Dim Dr, UDr%, U%, V, L%, J%
    U = -1
    For Each Dr In Dry
        UDr = UB(Dr)
        If UDr > U Then ReDim Preserve O(UDr)
        If Not IsEmptyAy(Dr) Then
            J = 0
            For Each V In Dr
                If IsNull(V) Then
                    L = 0
                Else
                    L = Len(V)
                End If
                If L > O(J) Then O(J) = L
                J = J + 1
            Next
        End If
    Next
DryWdtAy = O
End Function
Private Function DrLin$(Dr, Wdt%())
Dim UDr%
    UDr = UB(Dr)
Dim O$()
    Dim U%
    U = UB(Dr)
    ReDim O(U)
    Dim W, V
    Dim J%
    J = 0
    For Each W In Wdt
        V = ""
        If UDr >= J Then V = Dr(J)
        O(J) = AlignL(V, W)
        J = J + 1
    Next
DrLin = Quote(Join(O, " | "), "| * |")
End Function

