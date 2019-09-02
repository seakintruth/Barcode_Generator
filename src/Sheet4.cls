VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private mFontUpdating As Boolean

Private Sub frmTemplateOptions_Click()
    Debug.Print "Click", GetTemplateOption
End Sub

Private Sub frmTemplateOptions_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Debug.Print "KeyPress", GetTemplateOption
End Sub

Private Sub frmTemplateOptions_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Debug.Print "KeyUp", GetTemplateOption
End Sub

Private Sub Worksheet_Activate()
    ActionForWorkbook_Open
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    If Not mFontUpdating Then
        If Len(gstrFontNameCellName) > 0 Then
            Dim rngFont As Range: Set rngFont = Target.Parent.Range(gstrFontNameCellName)
            Dim rngFontSize As Range: Set rngFontSize = Target.Parent.Range(gstrFontSizeCellName)
            Debug.Print "Worksheet_Change", Now()
            Select Case Target.Address
                Case rngFont.Address
                    mFontUpdating = True
                    If IsValidFontName(rngFont.Value) Then
                        Union(rngFont, GetNonBlankCellsFromRange(Target.Parent.Columns(1))).Font.Name = rngFont.Value
                    End If
                    mFontUpdating = False
                Case rngFontSize.Address
                    mFontUpdating = True
                    GetNonBlankCellsFromRange(Target.Parent.Columns(1)).Font.Size = Target.Parent.Range(gstrFontSizeCellName).Value
                    mFontUpdating = False
            End Select
        End If
    End If
End Sub
