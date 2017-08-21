VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Log"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Sub Brw()
Ft.Brw
End Sub

Sub Log(Msg$, Optional FilNum%)
Dim F%
    F = FilNum
'    If FilNum = 0 Then F = LogFilNum
Print #F, NowStr & " " & Msg
If FilNum = 0 Then Close #F
End Sub

Property Get FilNum%()
'LogFilNum = Ft.OpnApp
End Property
Function Ft() As Ft
Set Ft = Nw.Ft(Pth & "Log.txt")
End Function

Function Pth$()
Dim O$: O = WrkPth & "Log\"
'PthEns O
Pth = O
End Function

