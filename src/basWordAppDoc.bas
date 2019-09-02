Attribute VB_Name = "basWordAppDoc"
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
'Word constants
Private Enum WordConstants
    cwdMainTextStory = 1
    cwdStory = 6
    cwdFormatOriginalFormatting = 16
    cwdFieldIncludePicture = 67
    cwdDialogInsertPicture = 163
End Enum

Public Sub ExportToWordDocument( _
        ByRef rngData As Range, _
        ByRef shtData As Worksheet, _
        Optional ByVal strFilePath As String, _
        Optional ByVal lngCellsPerPage As Long _
)
'We use Late binding for every word object so that we don't have to attach a referance to 'Microsoft Word xx.0 Object Library'
Dim wrd As Object
Dim docTemplate As Object
    Set wrd = GetOfficeApplication(oatcWord)
    wrd.Visible = True
    wrd.ScreenUpdating = False
    If Len(strFilePath & vbNullString) > 0 Then
        Set docTemplate = wrd.Documents.Open(strFilePath)
        'Set the start cell #
        Dim lngStartCell As Long
        If IsNumeric(shtData.Range(gstrTemplateStartCellName).Value) Then
            If lngStartCell > 0 And lngStartCell <= lngCellsPerPage Then
                lngStartCell = CInt(shtData.Range(gstrTemplateStartCellName).Value)
            Else
                lngStartCell = 1
            End If
        Else
            lngStartCell = 1
        End If
        'Count the number of barcodes to generate
        Dim lngLastBarcodeRow As Long
        'lngLastBarcodeRow = lngStartCell
        lngLastBarcodeRow = rngData.End(xlDown).Row
        'Generate the appropriate number of pages
        Dim intPages As Integer
        intPages = 0
        If ((lngLastBarcodeRow / lngCellsPerPage) - Int(lngLastBarcodeRow / lngCellsPerPage)) Then
            intPages = 1
        End If
        intPages = intPages + Int(lngLastBarcodeRow / lngCellsPerPage)
        'copy page
        docTemplate.Activate
        Dim rng As Object
        Set rng = docTemplate.Range
        '.WholeStory
        rng.Copy
        wrd.Selection.WholeStory
        wrd.Selection.Copy
        wrd.Selection.EndKey Unit:=cwdStory
        If intPages <> 1 Then
            Do Until intPages = 1 ' we start with 1 page, so we only need to add the extras
                intPages = intPages - 1
                wrd.Selection.PasteAndFormat cwdFormatOriginalFormatting
                'rng.PasteAndFormat (cwdFormatOriginalFormatting)
            Loop
            rng.Select 'TODO: should be able to do everything without using select
            wrd.Selection.EndKey Unit:=cwdStory 'end of range
            wrd.Selection.TypeBackspace 'removes the last blank page
        End If
        Dim lngBarcodeItem As Long
        lngBarcodeItem = lngStartCell
        Dim wrdTbl As Object
        Dim intCurrentWordPage As Integer
        intCurrentWordPage = 1
        Set wrdTbl = docTemplate.Tables(intCurrentWordPage)
        mRemoveTableCellPadding wrdTbl, wrd
        Dim intCurrentWordCellColumn As Integer
        intCurrentWordCellColumn = 1
        Dim intCurrentWordCellRow As Integer
        intCurrentWordCellRow = lngStartCell
        Dim wrdRngCell As Object
        Dim c As Range 'excel cell
        For Each c In rngData.Cells 'Excel Barcode data
            DoEvents
            Dim strImageDownloadFile As String
            strImageDownloadFile = DownloadUriFileToTemp(GetBarcodeImageUrl(c.Value), "png")
            If FileExists(strImageDownloadFile) Then
                lngBarcodeItem = c.Row + lngStartCell - 1
                Application.StatusBar = "Building Barcode " & lngBarcodeItem & " of " & lngLastBarcodeRow
                wrd.StatusBar = "Building Barcode " & lngBarcodeItem & " of " & lngLastBarcodeRow
                wrd.ScreenRefresh
                If intCurrentWordPage <> Int((lngBarcodeItem - 1) / lngCellsPerPage) + 1 Then
                    intCurrentWordPage = Int((lngBarcodeItem - 1) / lngCellsPerPage) + 1
                    Set wrdTbl = docTemplate.Tables(intCurrentWordPage)
                    mRemoveTableCellPadding wrdTbl, wrd
                End If
                'Set target cell address
                'This structrue is for the 30/page template
                Select Case shtData.Range(gstrTemplateCellName).Value
                    Case gOptTemplate.gOptAvery5167
                        intCurrentWordCellColumn = (((lngBarcodeItem - 1) Mod 4) * 2) + 1
                        intCurrentWordCellRow = Int((lngBarcodeItem - ((intCurrentWordPage - 1) * lngCellsPerPage) + 3) / 4)
                    Case gOptTemplate.gOptAvery5160
                        intCurrentWordCellColumn = (((lngBarcodeItem - 1) Mod 3) * 2) + 1
                        intCurrentWordCellRow = Int((lngBarcodeItem - ((intCurrentWordPage - 1) * lngCellsPerPage) + 2) / 3)
                End Select
                Set wrdRngCell = wrdTbl.Cell(intCurrentWordCellRow, intCurrentWordCellColumn)
                mInsertUrlGraphicInline strImageDownloadFile, wrdRngCell.Range.Characters.Last ' GetBarcodeImageUrl(c.Value), wrdRngCell 'Range to update docWord.Shapes(c.Row + lngStartCell)
            End If
        Next
    Else
    
        Set docTemplate = wrd.Documents.Add
        'Dump barcode images to the file
        Dim img As Object
        For Each c In rngData
            mInsertUrlGraphicInline GetBarcodeImageUrl(c.Value), docTemplate.StoryRanges(cwdMainTextStory).Characters.Last
            'Add some static space after each image...
            Dim gintSpacesAfterBarcode As Integer
            gintSpacesAfterBarcode = 5
            docTemplate.StoryRanges(cwdMainTextStory).Characters.Last.InsertAfter String(gintSpacesAfterBarcode, " ")
            
        Next
    End If
    wrd.ScreenUpdating = True
    wrd.Visible = True
    wrd.Activate
    SetEcho True, True
End Sub

Private Sub mRemoveTableCellPadding(ByRef tbl As Object, ByRef wrdApp As Object)

'    With tbl
'        .TopPadding = wrdApp.InchesToPoints(0)
'        .BottomPadding = wrdApp.InchesToPoints(0)
'        .LeftPadding = wrdApp.InchesToPoints(0)
'        .RightPadding = wrdApp.InchesToPoints(0)
''        .Spacing = 0
''        .AllowPageBreaks = True
''        .AllowAutoFit = False
'    End With
End Sub

Private Sub mInsertUrlGraphicInline(ByRef strUrlImage As String, ByRef rng As Object)
    'Strip HTML
    strUrlImage = Replace(strUrlImage, "<img src=""", "")
    strUrlImage = Replace(strUrlImage, """ border=""0"">", "")
    On Error Resume Next
    rng.InlineShapes.AddPicture strUrlImage, True, True
    If Err.Number <> 0 Then
        Err.Clear
        'try again once
        On Error GoTo 0
        rng.InlineShapes.AddPicture strUrlImage, True, True
    End If
    On Error GoTo 0
End Sub



