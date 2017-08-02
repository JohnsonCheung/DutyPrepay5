Attribute VB_Name = "bb_Lib_Ado_Flds"
Option Compare Database
Function AFldsDr(AFlds As ADODB.Fields) As Variant()
Dim O()
ReDim O(AFlds.Count - 1)
Dim J%, F As ADODB.Field
For Each F In AFlds
    O(J) = F.Value
    J = J + 1
Next
AFldsDr = O
End Function
Function AFldsFny(AFlds As ADODB.Fields) As String()
Dim O$()
Dim F As ADODB.Field
For Each F In AFlds
    Push O, F.Name
Next
AFldsFny = O
End Function
