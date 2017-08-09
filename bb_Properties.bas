Attribute VB_Name = "bb_Properties"
Option Compare Database
Option Explicit

Function CurFb$()
CurFb = CurrentDb.Name
End Function

Function CurPth$()
CurPth = FfnPth(CurFb)
End Function

Function DtaDb() As Dao.Database
Set DtaDb = DBEngine.OpenDatabase(DtaFb)
End Function

Function DtaFb$()
DtaFb = FfnRplExt(FfnAddFnSfx(CurFb, "_Data"), ".mdb")
End Function

Function PermitImpPth$()
Dim O$: O = CurPth & "Import - Permit\"
PthEns O
PermitImpPth = O
End Function

Function WrkPth$()
WrkPth = CurPth & "WorkingDir\"
End Function
