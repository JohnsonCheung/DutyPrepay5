Attribute VB_Name = "Dao_Prp"
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

Function WrkPth$()
WrkPth = CurPth & "WorkingDir\"
End Function
