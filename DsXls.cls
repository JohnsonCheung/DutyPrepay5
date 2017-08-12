VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DsXls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Private A As Ds
Property Get Ds() As Ds
Set Ds = A
End Property
Function Init(Ds As Ds) As DsXls
If IsNothing(A) Then Stop
Set A = Ds
Set Init = Me
End Function

Function Wb() As Workbook
Dim O As Workbook
Set O = WbNew
With WbFstWs(O)
    .Name = "Ds"
    .Range("A1").Value = A.DsNm
End With
Dim Dt As Dt, I
If Not Ds.IsEmpty Then
    For Each I In Ds.DtAy
        Set Dt = I
        Dt.Xls.AddWb O
    Next
End If
Set Wb = O
End Function

Private Sub Wb__Tst()
Dim Wb As Workbook
Set Wb = Db.Ds("Permit PermitD").Xls.Wb
WbVis Wb
Stop
Wb.Close False
End Sub

