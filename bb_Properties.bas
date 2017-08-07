Attribute VB_Name = "bb_Properties"
Option Compare Database

Property Get CurFb$()
CurFb = CurrentDb.Name
End Property

Property Get CurPth$()
CurPth = FfnPth(CurFb)
End Property

Property Get DtaDb() As DAO.Database
Set DtaDb = DBEngine.OpenDatabase(DtaFb)
End Property

Property Get DtaFb$()
DtaFb = AddFnSfx(CurFb, "_Data")
End Property

Function PermitImpPth$()
Dim O$: O = CurPth & "Import - Permit\"
PthEns O
PermitImpPth = O
End Function

Property Get WrkPth$()
WrkPth = CurPth & "WorkingDir\"
End Property
