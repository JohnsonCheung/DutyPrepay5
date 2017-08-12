VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SqXls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Private Sq As Sq
Function Init(A As Sq) As SqXls
Set Sq = A
Set Init = Me
End Function

Sub PutAt(At As Range)
ReSzRg(At, Sq).Value = Sq
End Sub
