VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AFlds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Public Flds As ADODB.Fields
Function Init(A As ADODB.Fields) As AFlds
Set Flds = A
Set Init = Me
End Function
Function Dr() As Variant()
Dim O()
ReDim O(Flds.Count - 1)
Dim J%, F As ADODB.Field
For Each F In Flds
    O(J) = F.Value
    J = J + 1
Next
Dr = O
End Function
Function DrX() As Ay
Set DrX = Ay(Dr)
End Function
Function Fny() As String()
Dim O$()
Dim F As ADODB.Field
For Each F In Flds
    Push O, F.Name
Next
Fny = O
End Function

Function FnyX() As Ay
Set FnyX = Ay(Fny)
End Function
