VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database

Function Db() As Db
Set Db = Nw.Db(DBEngine.CreateDatabase(Fb, Dao.LanguageConstants.dbLangGeneral))
End Function

Function Fb$(Optional Fdr$)
Fb = Pth(Fdr) & Fn(".accdb")
End Function

Function Fn$(Ext$)
Fn = Nm & Ext
End Function

Function Ft$(Optional Fdr$)
Ft = Pth(Fdr) & Fn(".txt")
End Function

Function Fx$(Optional Fdr$)
Fx = Pth(Fdr) & Fn(".xlsx")
End Function

Function Nm$()
Static X&
Nm = "T" & Format(Now(), "YYYYMMDD_HHMMSS") & "_" & X
X = X + 1
End Function

Function Pth$(Optional Fdr$)
Static X$
If X = "" Then X = Fso.GetSpecialFolder(TemporaryFolder) & "\"
If Fdr = "" Then
    Pth = X
Else
    Dim O$
    O = X & Fdr & "\"
    Nw.Pth(O).Ens
    Pth = O
End If
End Function
