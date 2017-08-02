VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmBF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Option Base 0
Private Sub CmdSetBusArea_Click()
DoCmd.OpenForm "frmBF"
End Sub
Private Sub Form_Load()
DoCmd.Maximize
Me.RecordSource = "SELECT x.BusArea AS CdBusArea, a.BusAreaName AS NmBusArea, x.BrandFamilyName AS CdBF, x.BrandFamilyDesc AS NmBF" & _
" FROM tblSAPBrandFamily AS x LEFT JOIN tblSAPBusArea AS a ON x.BusArea = a.BusArea" & _
" Where BrandFamilyDesc<>'0'" & _
" ORDER BY a.BusAreaName, x.BrandFamilyDesc;"
End Sub
