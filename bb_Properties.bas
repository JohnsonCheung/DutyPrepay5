Attribute VB_Name = "bb_Properties"
Option Compare Database
Property Get DtaFb$()
DtaFb = AddFnSfx(CurFb, "_Data")
End Property
Property Get CurFb$()
CurFb = CurrentDb.Name
End Property
Property Get DtaDb() As Dao.Database
Set DtaDb = DBEngine.OpenDatabase(DtaFb)
End Property
Property Get WrkPth$()
WrkPth = CurPth & "WorkingDir\"
End Property
Property Get CurPth$()
CurPth = FfnPth(CurFb)
End Property
Property Get PermitImpPth$()
Dim O$: O = CurPth & "Import - Permit\"
EnsPth O
PermitImpPth = O
End Property
