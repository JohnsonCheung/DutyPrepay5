Attribute VB_Name = "bb_Lib_Dta_Ins"
Option Compare Database
Option Explicit


Property Get SampleDt() As Dt
Dim Dry As New Dry: Dry.Push Array(1, 2, 3)
Dim Fny$(): Fny = SplitLvs("A B C")
Set SampleDt = Dt(Fny, Dry, "Sample")
End Property

