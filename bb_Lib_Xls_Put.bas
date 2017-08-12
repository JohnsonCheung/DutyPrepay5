Attribute VB_Name = "bb_Lib_Xls_Put"
Option Compare Database
Option Explicit

Sub AyPut(Ay, Cell As Range)
SqByHAy(Ay).Xls.PutAt Cell
End Sub
Sub DryPut(AtCell As Range, Dry As Dry)
AtCell.Value = Dry.Sq
End Sub

