Attribute VB_Name = "basInputActions"
Option Explicit
'Authored 2014-2019 by Jeremy Dean Gerdes <jeremy.gerdes@navy.mil>
    'Public Domain in the United States of America,
     'any international rights are waived through the CC0 1.0 Universal public domain dedication <https://creativecommons.org/publicdomain/zero/1.0/legalcode>
     'http://www.copyright.gov/title17/
     'In accrordance with 17 U.S.C. § 105 This work is 'noncopyright' or in the 'public domain'
         'Subject matter of copyright: United States Government works
         'protection under this title is not available for
         'any work of the United States Government, but the United States
         'Government is not precluded from receiving and holding copyrights
         'transferred to it by assignment, bequest, or otherwise.
     'as defined by 17 U.S.C § 101
         '...
         'A “work of the United States Government” is a work prepared by an
         'officer or employee of the United States Government as part of that
         'person’s official duties.
         '...

Public Sub DropDown9_Change()
    If Len(gstrBarcode3Of9) = 0 Then
        SetGlobalVariables
    End If
    Dim rngBarcode3Of9 As Range
    Set rngBarcode3Of9 = ThisWorkbook.Sheets(gstrOutputSheetName).Range(gstrBarcode3Of9)
End Sub

Public Function GetSelectedTemplateOption()
    'gSelectedTemplateNumber is only ever set after we use GetSelectedTemplateOption
    Dim sht As Worksheet
    If Len(gstrDataSheetName) = 0 Then
        SetGlobalVariables
    End If
    Set sht = ThisWorkbook.Sheets(gstrDataSheetName)
    Dim ctl As Control
    For Each ctl In sht.OLEObjects("frmTemplateOptions").Object.Controls
        If ctl.Value Then
            gSelectedTemplateNumber = ctl.TabIndex
            Select Case ctl.TabIndex
                Case gOptAvery5167
                    GetSelectedTemplateOption = gstrOptionAvery5167Name
                Case gOptAvery5160
                    GetSelectedTemplateOption = gstrOptionAvery5160Name
                Case gOptAvery5262
                    GetSelectedTemplateOption = gstrOptionAvery5262Name
                Case gOptAvery5360
                    GetSelectedTemplateOption = gstrOptionAvery5360Name
                Case gOptCustom
                    GetSelectedTemplateOption = gstrOptionCustomName
            End Select
        End If
    Next ctl
End Function
