Attribute VB_Name = "bb_Lib_Vb_Log"
Option Compare Database
Option Explicit
Sub Log(Msg$, Optional FilNum%)
Dim F%
    F = FilNum
    If FilNum = 0 Then F = LogFilNum
Print #F, NowStr & " " & Msg
If FilNum = 0 Then Close #F
End Sub
Property Get LogFilNum%()
LogFilNum = OpnApp(LogFt)
End Property
Property Get LogFt$()
LogFt = LogPth & "Log.txt"
End Property
Property Get LogPth$()
Dim O$: O = WrkPth & "Log\"
EnsPth O
LogPth = O
End Property
Sub BrwLog()
BrwFt LogFt
End Sub
