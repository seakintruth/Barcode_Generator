VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Activate()
    SetGlobalVariables
    ActiveSheet.Visible = xlSheetVeryHidden
    Dim sht As Worksheet
    Set sht = ThisWorkbook.Worksheets(gstrDataSheetName)
    sht.Visible = xlSheetVisible
    sht.Activate
End Sub
