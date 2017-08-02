Attribute VB_Name = "bb_Lib_Dao_Flds"
Option Compare Database
Option Explicit
Function FldsDr(Flds As Dao.Fields) As Variant()
Dim O()
ReDim O(Flds.Count - 1)
Dim J%, F As Dao.Field
For Each F In Flds
    O(J) = F.Value
    J = J + 1
Next
FldsDr = O
End Function

