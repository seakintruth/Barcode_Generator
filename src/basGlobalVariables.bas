Attribute VB_Name = "basGlobalVariables"
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
'----------------------------------------------------------------------------
'Validation Tests should be (not yet written/implemented):
    'After calling SetGlobalVariables (these would be code implementation tests, not behaviour tests):
        'Each Global Variable in this Module should not be null/empty
        'Each Variable fom the worksheet 'StaticValues' column C should have a value set as a global variable
'----------------------------------------------------------------------------
'Global Variables
Global Const gstrOptionsSheetName  As String = "StaticValues"
Global Const gstrNoticeSheetName  As String = "Notice"
Public gstrDataSheetName As String
Public gstrOutputSheetName As String
Public gstrCustomRangeColumns  As String
Public gstrTextBoxHeightCellName As String

'Fixed Cells
Public gstrBarcode3Of9 As String
Public gstrTemplateCellName  As String
'Public gstrStyleCellName As String
'Public gstrTypeCellName  As String
Public gstrWidthScaleCellName  As String
Public gstrColumnsOfBarcodes As String
Public gstrHeightCellName  As String
Public gstrFontSizeCellName As String
Public gstrTextPositionCellName As String
Public gstrFontNameCellName As String
Public gstrBarcodeResolutionCellName As String
Public gstrLeftPaddingImageCellName As String
Public gstrRightPaddingImageCellName As String
Public gstrTopPaddingImageCellName As String
Public gstrBottomPaddingImageCellName As String
Public gstrDestinationTypeCellName As String
Public gstrTemplateStartCellName  As String
Public gstrSymbologyCellName As String

'Template Download URLs
Public gstrAveryTemplatePath As String

'Template Names
Public gstrOptionAvery5167Name  As String
Public gstrOptionAvery5160Name  As String
Public gstrOptionAvery5360Name  As String
Public gstrOptionAvery5262Name  As String
Public gstrOptionAvery5167FileName  As String
Public gstrOptionAvery5160FileName  As String
Public gstrOptionAvery5360FileName  As String
Public gstrOptionAvery5262FileName  As String
Public gstrOptionCustomName  As String

'Display Names
Public gstrOptionStyleDisplayText  As String
Public gstrOptionStyleStretch  As String
Public gstrOptionStyleNoText  As String

'Barcode Attributes
Public gLngBarcodeWidth As Long
Public gLngBarcodeHeight As Long
Public gLngBarcodeFontSize As Long
Public gLngBarcodeResolution As Long
Public gLngOptionTemplateValue As Long
Public gLngOptionStyleValue As Long
Public glngStyleValue As Long
'Public gstrType As String

Public Enum gOptTemplate
    gOptAvery5167 = 0
    gOptAvery5160 = 1
    gOptAvery5360 = 2
    gOptAvery5262 = 3
    gOptCustom = 4
End Enum

Public gSelectedTemplateNumber As Integer

Public Enum gOptDestinationType
    gOptBlankWordDoc = 1
    gOptHTML = 2
    gOptAveryTemplate = 3
End Enum

Public Enum gOptStyle
    gOptStyleDisplayText = 2
    gOptStyleStretchText = 3
    gOptStyleNoText = 1
End Enum

Public Sub GetStaticValue(ByVal strVariableName As String, ByRef strValueToSet As String, Optional ByRef rngSource As Range)
If IsNothing(rngSource) Then
    Set rngSource = Intersect(ThisWorkbook.Sheets("StaticValues").Columns(3), ThisWorkbook.Sheets("StaticValues").UsedRange)
End If
Dim c As Range
    For Each c In rngSource
        If Len(strValueToSet) = 0 Then
            If Trim(LCase(c.Value)) = Trim(LCase(strVariableName)) Then
                strValueToSet = c.Offset(0, -1)
                GoTo ExitHere
            End If
        End If
    Next
ExitHere:
    Set c = Nothing
End Sub

Public Sub SetGlobalVariables()
'------------------------
Dim sht As Worksheet: Set sht = ThisWorkbook.Sheets("StaticValues")
'Setting rngSource prior to calling GetStaticValue in an effor to speed up this action
Dim rngSource As Range: Set rngSource = Intersect(sht.Columns(3), sht.UsedRange)

    '[Ranges]
    GetStaticValue "gstrBarcode3Of9", gstrBarcode3Of9, rngSource
    GetStaticValue "gstrOutputSheetName", gstrOutputSheetName, rngSource
    GetStaticValue "gstrDataSheetName", gstrDataSheetName, rngSource
    GetStaticValue "gstrTemplateCellName", gstrTemplateCellName, rngSource
    GetStaticValue "gstrDataSheetName", gstrDataSheetName, rngSource
    GetStaticValue "gstrColumnsOfBarcodes", gstrColumnsOfBarcodes, rngSource
    GetStaticValue "gstrCustomRangeColumns", gstrCustomRangeColumns, rngSource
'    GetStaticValue "gstrStyleCellName", gstrStyleCellName, rngSource
'    GetStaticValue "gstrTypeCellName", gstrTypeCellName, rngSource
    GetStaticValue "gstrWidthScaleCellName", gstrWidthScaleCellName, rngSource
    GetStaticValue "gstrHeightCellName", gstrHeightCellName, rngSource
    GetStaticValue "gstrFontSizeCellName", gstrFontSizeCellName, rngSource
    GetStaticValue "gstrFontNameCellName", gstrFontNameCellName, rngSource
    GetStaticValue "gstrTextPositionCellName", gstrTextPositionCellName, rngSource
    GetStaticValue "gstrTextBoxHeightCellName", gstrTextBoxHeightCellName, rngSource
    GetStaticValue "gstrBarcodeResolutionCellName", gstrBarcodeResolutionCellName, rngSource
    GetStaticValue "gstrLeftPaddingImageCellName", gstrLeftPaddingImageCellName, rngSource
    GetStaticValue "gstrRightPaddingImageCellName", gstrRightPaddingImageCellName, rngSource
    GetStaticValue "gstrTopPaddingImageCellName", gstrTopPaddingImageCellName, rngSource
    GetStaticValue "gstrBottomPaddingImageCellName", gstrBottomPaddingImageCellName, rngSource
    GetStaticValue "gstrDestinationTypeCellName", gstrDestinationTypeCellName, rngSource
    GetStaticValue "gstrTemplateStartCellName", gstrTemplateStartCellName, rngSource
    GetStaticValue "gstrSymbologyCellName", gstrSymbologyCellName, rngSource
    '[TemplateNames]
    GetStaticValue "gstrOptionAvery5167Name", gstrOptionAvery5167Name, rngSource
    GetStaticValue "gstrOptionAvery5160Name", gstrOptionAvery5160Name, rngSource
    GetStaticValue "gstrOptionAvery5360Name", gstrOptionAvery5360Name, rngSource
    GetStaticValue "gstrOptionAvery5262Name", gstrOptionAvery5262Name, rngSource
    GetStaticValue "gstrOptionAvery5167FileName", gstrOptionAvery5167FileName, rngSource
    GetStaticValue "gstrOptionAvery5160FileName", gstrOptionAvery5160FileName, rngSource
    GetStaticValue "gstrOptionAvery5360FileName", gstrOptionAvery5360FileName, rngSource
    GetStaticValue "gstrOptionAvery5262FileName", gstrOptionAvery5262FileName, rngSource
    GetStaticValue "gstrOptionCustomName", gstrOptionCustomName, rngSource
    
    '[ServiceProviderNames]
    GetStaticValue "gstrOptionStyleDisplayText", gstrOptionStyleDisplayText, rngSource
    GetStaticValue "gstrOptionStyleStretch", gstrOptionStyleStretch, rngSource
    GetStaticValue "gstrOptionStyleNoText", gstrOptionStyleNoText, rngSource
    
    '[TemplateDownloadURLs]
    GetStaticValue "gstrAveryTemplatePath", gstrAveryTemplatePath, rngSource
    'GetStaticValue "gstrPhpBarcodeGeneratorServiceProvider", gstrPhpBarcodeGeneratorServiceProvider, rngSource
    'GetStaticValue "gstrPhpAlternateBarcodeGeneratorServiceProvider", gstrPhpAlternateBarcodeGeneratorServiceProvider, rngSource
    
End Sub
