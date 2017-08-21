VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PjNwCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Public Pj_ As Pj
Function Init(Pj As Pj) As PjNwCls
Set Pj_ = Pj
Set Init = Me
End Function
Sub Upd()
'Each class has function with pfx-Init.  Use it to update function in module-Nw
Nw_RmvFun
Nw_AddFun
End Sub
Private Sub Nw_RmvFun()

End Sub

Private Sub Nw_AddFun()

End Sub
