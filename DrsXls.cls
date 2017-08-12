VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DrsXls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Public Drs As Drs
Function Init(A As Drs) As Drs
If IsNothing(A) Then Stop
Set Drs = A
Set Init = Me
End Function
Function Ws(Optional WsNm$ = "Sheet1", Optional Vis As Boolean) As Worksheet
Dim O As Worksheet: Set O = WsNew(WsNm, Vis:=True)
PutAt WsA1(O)
If Vis Then WsVis O
Set Ws = O
End Function

Sub PutAt(At As Range, Optional LoNm$)
AyPut Drs.Fny, At
Drs.Dry.Sq.Xls.PutAt RgRC(At, 2, 1)
LoCrt RgWs(At), LoNm
End Sub



