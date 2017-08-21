VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Flds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Public Flds As Dao.Fields
Function Init(A As Dao.Fields) As Flds
Set Flds = A
Set Init = Me
End Function
Function Fny() As String()
Dim O$()
Dim F As Dao.Field
For Each F In Flds
    Push O, F.Name
Next
Fny = O
End Function

Function HasFld(F) As Boolean
Dim I As Dao.Field
For Each I In Flds
    If I.Name = F Then HasFld = True: Exit Function
Next
End Function

Function Dr() As Variant()
Dim O()
ReDim O(Flds.Count - 1)
Dim J%, F As Dao.Field
For Each F In Flds
    O(J) = F.Value
    J = J + 1
Next
Dr = O
End Function

Function NFld%()
NFld = Flds.Count
End Function

Function InsSqlValPhraseTp$()
Dim F As Dao.Field, O$()
For Each F In Flds
    Push O, Nw.Fld(F).InsSqlValPhraseTp
Next
InsSqlValPhraseTp = JnComma(O)
End Function
