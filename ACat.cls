VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ACat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Public Cat_ As Catalog

Function Init(Cat As Catalog) As ACat
Set Cat_ = Cat
End Function
Function Tny() As String()
Dim T As ADOX.Table
Dim O$()
For Each T In Cat_.Tables
    Push O, T.Name
Next
Tny = O
End Function
Function Fny(T) As String()
Dim C As ADOX.Column
Dim O$()
For Each C In Cat_.Tables(T).Columns
    Push O, C.Name
Next
Fny = O
End Function
Function FnyX(T) As Ay
Set FnyX = Ay(Fny(T))
End Function
Function TnyX() As Ay
Set TnyX = Ay(Tny)
End Function
