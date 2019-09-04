Attribute VB_Name = "basUserActions"
Option Explicit
'Authored 2014-2017 by Jeremy Dean Gerdes <jeremy.gerdes@navy.mil>
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

Sub grpBoxTemplates_Click()
SetEcho False
If Len(gstrDataSheetName) = 0 Then
    SetGlobalVariables
End If
Dim sht As Worksheet: Set sht = Worksheets(gstrDataSheetName)
gLngOptionTemplateValue = sht.Range(gstrTemplateCellName).Value
    Select Case gLngOptionTemplateValue
        Case gOptAvery5167
            gLngBarcodeWidth = 190
            gLngBarcodeHeight = 45
            gLngBarcodeFontSize = 1
            gLngBarcodeResolution = 1
            mSetBarcodeAttributes sht
            mHideCustomEntryColumns
            sht.Range(gstrTemplateCellName).Offset(0, 1).Value _
                = GetTemplateOption(sht.Range(gstrTemplateCellName).Value)
        Case gOptAvery5160
            gLngBarcodeWidth = 250
            gLngBarcodeHeight = 92
            gLngBarcodeFontSize = 3
            gLngBarcodeResolution = 1
            mSetBarcodeAttributes sht
            mHideCustomEntryColumns
            sht.Range(gstrTemplateCellName).Offset(0, 1).Value _
                = GetTemplateOption(sht.Range(gstrTemplateCellName).Value)
        Case gOptAvery5360
            gLngBarcodeWidth = 250
            gLngBarcodeHeight = 92
            gLngBarcodeFontSize = 3
            gLngBarcodeResolution = 1
            mSetBarcodeAttributes sht
            mHideCustomEntryColumns
            sht.Range(gstrTemplateCellName).Offset(0, 1).Value _
                = GetTemplateOption(sht.Range(gstrTemplateCellName).Value)
        Case gOptAvery5262
            gLngBarcodeWidth = 250
            gLngBarcodeHeight = 92
            gLngBarcodeFontSize = 3
            gLngBarcodeResolution = 1
            mSetBarcodeAttributes sht
            mHideCustomEntryColumns
            sht.Range(gstrTemplateCellName).Offset(0, 1).Value _
                = GetTemplateOption(sht.Range(gstrTemplateCellName).Value)
        Case gOptCustom
            mGetBarcodeAttributes sht
            If sht.Range(gstrDestinationTypeCellName).Value = 3 Then
                MsgBox _
                    "Barcode template can not be 'Custom' and have a " & _
                    "destination of 'Avery word Template'", vbOKOnly, "Barcode Generator"
                sht.Range(gstrDestinationTypeCellName).Value = 2
            End If
            mShowCustomEntryColumns
            sht.Range(gstrTemplateCellName).Offset(0, 1).ClearContents
    End Select
SetEcho True
End Sub

Sub grpBoxStyle_Click()
SetEcho False
If Len(gstrDataSheetName) = 0 Then
    SetGlobalVariables
    If Len(gstrDataSheetName) = 0 Then
        MsgBox "Missing Additional Option Value(s):" & vbCrLf & vbCrLf & "Data Sheet Name or Style Cell Name", vbCritical + vbOKOnly, "Barcode Generator"
    End If
End If
Dim sht As Worksheet: Set sht = Worksheets(gstrDataSheetName)
'gLngOptionTemplateValue = sht.Range(gstrStyleCellName).Value
    Select Case gLngOptionTemplateValue
        Case gOptStyleDisplayText
            glngStyleValue = 196
        Case gOptStyleStretchText
            glngStyleValue = 452
        Case gOptStyleNoText
            glngStyleValue = 68
    End Select
SetEcho True
End Sub

Sub optBtnDestinationAvery_Click()
Dim wsh As Worksheet
    Set wsh = ThisWorkbook.ActiveSheet
    If Len(gstrTemplateCellName) = 0 Then
        SetGlobalVariables
        If Len(gstrTemplateCellName) = 0 Then
            MsgBox "Missing Additional Option Value(s):" & vbCrLf & vbCrLf & "Template Cell Name", vbCritical + vbOKOnly, "Barcode Generator"
        End If
    End If
    If wsh.Range(gstrTemplateCellName).Value = 3 Then
        MsgBox _
            "Barcode template can not be 'Custom' and have a " & _
            "destination of 'Avery word Template'", vbOKOnly, "Barcode Generator"
        wsh.Range(gstrDestinationTypeCellName).Value = 2
    End If
End Sub

Sub btnGotoAdditionalOptions_Click()
    SetCustomAppOptions
        Dim sht As Worksheet
        Set sht = ThisWorkbook.Sheets(gstrOptionsSheetName)
        sht.Visible = xlSheetVisible
        sht.Activate
    SetOriginalAppOptions
End Sub

Sub btnExportBarcodeToFile_Click()
'[TODO] Prior to re-implementing output to HTML, or Word templates, need to capture the barcode object as an image
'Generate Barcode File
SetEcho False
''''''''''''''''''''''''''''''''''''''''''''
'inputs validation should be added here
''''''''''''''''''''''''''''''''''''''''''''
    SetGlobalVariables

    Dim shtData As Worksheet: Set shtData = Worksheets(gstrDataSheetName)
    Dim rngData As Range
    Set rngData = Intersect(GetNonBlankCellsFromWorksheet(shtData), shtData.Range("A:A"))
    If IsNothing(rngData) Then
    'If Evaluate("Counta(A:A)") = 0 Then
        MsgBox "You must enter data in column A to generate barcodes", vbInformation, "Barcode Generator, Need User Input"
    Else
        mGetBarcodeAttributes shtData
        Select Case shtData.Range(gstrDestinationTypeCellName)
            Case gOptDestinationType.gOptAveryTemplate
                'Download the file if it's not found in the barcode generator's Tempates directory
                Dim strTemplateFilePath As String
                Dim strCurrentDir As String
                Dim lngCellsPerPage As Long
                strCurrentDir = ThisWorkbook.path
                Select Case shtData.Range(gstrTemplateCellName).Value
                    Case gOptTemplate.gOptAvery5167
                        strTemplateFilePath = mPickOrGetTempateFile(strCurrentDir, gstrOptionAvery5167FileName, gstrAveryTemplatePath)
                        lngCellsPerPage = 80
                    Case gOptTemplate.gOptAvery5360
                        strTemplateFilePath = mPickOrGetTempateFile(strCurrentDir, gstrOptionAvery5360FileName, gstrAveryTemplatePath)
                        lngCellsPerPage = 21
                    Case gOptTemplate.gOptAvery5262
                        strTemplateFilePath = mPickOrGetTempateFile(strCurrentDir, gstrOptionAvery5262FileName, gstrAveryTemplatePath)
                        lngCellsPerPage = 14
                    Case gOptTemplate.gOptAvery5160
                        strTemplateFilePath = mPickOrGetTempateFile(strCurrentDir, gstrOptionAvery5160FileName, gstrAveryTemplatePath)
                        lngCellsPerPage = 30
                End Select
                ExportToWordDocument rngData, shtData, strTemplateFilePath, lngCellsPerPage
            Case gOptDestinationType.gOptBlankWordDoc
                ExportToWordDocument rngData, shtData
            Case gOptDestinationType.gOptHTML
                Dim c As Range
                Dim Folder
                Dim tf ' As Scripting.TextStream
                Dim FSO As Object: Set FSO = CreateObject("Scripting.FileSystemObject") ' New Scripting.FileSystemObject
                Dim strExportFile As String: strExportFile = Environ("TEMP") & "\" & _
                    "Barcode Generator.html"
                If FSO.FileExists(strExportFile) Then
                    FSO.DeleteFile strExportFile
                End If
                Set tf = FSO.CreateTextFile(strExportFile, True)
                tf.WriteLine "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01//EN"">" & _
                    "<html><head><title> Template Generated Barcodes from Barcodesinc " & _
                    "by Gerdes, Jeremy D.</title></head><body>"
                For Each c In rngData
                    'tf.WriteLine GetBarcodeImageUrl(c.Value)
                    '[TODO] Prior to re-implementing output to HTML, or Word templates, need to capture the barcode object as an image
                Next
                tf.WriteLine "</body></html>"
                tf.Close
                OpenFileInChrome strExportFile
                'Cleanup
                Set tf = Nothing
                Set c = Nothing
                shtData.Activate
        End Select
    End If
    Set rngData = Nothing
    Set shtData = Nothing
    SetEcho True
End Sub

Private Function mPickOrGetTempateFile( _
    strCurrentDir As String, _
    strFileName As String, _
    strFilePath As String _
) As String
    Dim strTemplateFilePath As String
    If FolderExists(strFilePath) Then
        If Not FolderExists(strCurrentDir & "\Templates") Then
            If BuildDir(strCurrentDir & "\" & "Templates") Then
                CopyFolder strFilePath, strCurrentDir & "\" & "Templates"
            End If
        End If
        Select Case True
            Case FileExists(strCurrentDir & "\Templates\" & strFileName & ".dotx")
                strTemplateFilePath = strCurrentDir & "\Templates\" & strFileName & ".dotx"
            Case FileExists(strCurrentDir & "\Templates\" & strFileName & ".docx")
                strTemplateFilePath = strCurrentDir & "\Templates\" & strFileName & ".docx"
            Case FileExists(strCurrentDir & "\Templates\" & strFileName & ".doct")
                strTemplateFilePath = strCurrentDir & "\Templates\" & strFileName & ".doct"
            Case FileExists(strCurrentDir & "\Templates\" & strFileName & ".doc")
                strTemplateFilePath = strCurrentDir & "\Templates\" & strFileName & ".doc"
            Case Else
                If _
                    Not (FileExists(strFilePath & "\Templates\" & strFileName & ".dotx")) _
                    And Not (FileExists(strFilePath & "\Templates\" & strFileName & ".docx")) _
                    And Not (FileExists(strFilePath & "\Templates\" & strFileName & ".dot")) _
                    And Not (FileExists(strFilePath & "\Templates\" & strFileName & ".doc")) _
                Then
                    MsgBox "Unable to find required Template file at: " & vbCrLf & vbCrLf & strFilePath & "\Templates\" & strFileName & "...", vbCritical + vbOKOnly, "Barcode Generator"
                    strTemplateFilePath = vbNullString
                Else
                    If BuildDir(strCurrentDir & "\" & "Templates") Then
                        CopyFolder strFilePath, strCurrentDir & "\" & "Templates"
                    Else
                        'Resort to saving to temp folder, since we can't save in the current folder.
                        CopyFolder strFilePath, Environ("TEMP") & "\" & "Templates"
                        mPickOrGetTempateFile Environ("TEMP"), strFileName, strFilePath
                    End If
                    'DownloadTemplateToTemp returns a new strCurrentDir that is the parent path to a copy of 'Templates'
                    If LCase(Trim(strCurrentDir & "\Templates")) = LCase(Trim(gstrAveryTemplatePath)) Then
                        MsgBox "Unable to find expected file in Templates folder:" & vbCrLf & vbCrLf & gstrAveryTemplatePath, vbCritical + vbOKOnly, "Barcode Generator"
                        strTemplateFilePath = vbNullString
                    Else
                        strTemplateFilePath = mPickOrGetTempateFile(DownloadTemplateToTemp(strFilePath, strCurrentDir & "\" & "Templates"), strFileName, strFilePath)
                    End If
                End If
        End Select
    Else
        MsgBox "Unable to find expected Templates folder:" & vbCrLf & vbCrLf & strFilePath, vbCritical + vbOKOnly, "Barcode Generator"
        strTemplateFilePath = vbNullString
    End If
    mPickOrGetTempateFile = strTemplateFilePath
End Function

Private Sub mShowCustomEntryColumns()
    ThisWorkbook.Sheets(gstrDataSheetName).Columns(gstrCustomRangeColumns).EntireColumn.Hidden = False
End Sub

Private Sub mHideCustomEntryColumns()
    Worksheets(gstrDataSheetName).Columns(gstrCustomRangeColumns).EntireColumn.Hidden = True
End Sub

Private Sub mSetBarcodeAttributes(ByRef sht As Worksheet)
    sht.Range(gstrWidthScaleCellName).Value = gLngBarcodeWidth
    sht.Range(gstrHeightCellName).Value = gLngBarcodeHeight
    sht.Range(gstrFontSizeCellName).Value = gLngBarcodeFontSize
    sht.Range(gstrBarcodeResolutionCellName).Value = gLngBarcodeResolution
    sht.Range(gstrTemplateCellName).Value = gLngOptionTemplateValue
End Sub

Private Sub mGetBarcodeAttributes(ByRef sht As Worksheet)
    gLngOptionTemplateValue = sht.Range(gstrTemplateCellName).Value
    gLngBarcodeWidth = sht.Range(gstrWidthScaleCellName).Value
    gLngBarcodeHeight = sht.Range(gstrHeightCellName).Value
    gLngBarcodeFontSize = sht.Range(gstrFontSizeCellName).Value
    gLngBarcodeResolution = sht.Range(gstrBarcodeResolutionCellName).Value
'    gstrType = mGetSelectedTypeValue(sht)
End Sub
'
'Private Function mGetSelectedTypeValue(sht As Worksheet) As String
'    Select Case sht.Range(gstrTypeCellName).Value
'        Case 1
'            mGetSelectedTypeValue = "I25"
'        Case 2
'            mGetSelectedTypeValue = "C39"
'        Case 3
'            mGetSelectedTypeValue = "C128A"
'        Case 4
'            mGetSelectedTypeValue = "C128B"
'        Case 5
'            mGetSelectedTypeValue = "C128C"
'        Case Else
'            mGetSelectedTypeValue = "C128B"
'    End Select
'End Function

