Attribute VB_Name = "bb_Lib_Vb_Log"
Option Compare Database
Option Explicit

Sub BrwLog()
FtBrw LogFt
End Sub

Sub Log(Msg$, Optional FilNum%)
Dim F%
    F = FilNum
    If FilNum = 0 Then F = LogFilNum
Print #F, NowStr & " " & Msg
If FilNum = 0 Then Close #F
End Sub

Property Get LogFilNum%()
LogFilNum = FtOpnApp(LogFt)
End Property

Property Get LogFt$()
LogFt = LogPth & "Log.txt"
End Property

Property Get LogPth$()
Dim O$: O = WrkPth & "Log\"
PthEns O
LogPth = O
End Property
