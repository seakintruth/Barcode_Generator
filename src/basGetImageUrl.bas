Attribute VB_Name = "basGetImageUrl"
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


'Public Function GetBarcodeImageUrl(strValue As Variant)
'    If Len(strValue) <> 0 Then
'        Dim sht As Worksheet
'        Set sht = ThisWorkbook.Sheets(gstrDataSheetName)
'        Select Case True
'            Case gOptAvery5160 = gLngOptionTemplateValue
'                'GetBarcodeImageUrl = mGetBarcodeImageFromSheet(strValue, sht)
'            Case gOptAvery5167 = gLngOptionTemplateValue
'                'GetBarcodeImageUrl = mGetBarcodeImageFromSheet(strValue, sht)
'            Case gOptCustom = gLngOptionTemplateValue
'                'GetBarcodeImageUrl = mGetBarcodeImageFromSheet(strValue, sht)
'            Case Else
''                GetBarcodeImageUrl = mGetBarcodeImageUrl( _
''                    strValue, _
''                    , _
''                    , _
''                    , _
''                    glngStyleValue, _
''                    gstrType, _
''                    gLngBarcodeResolution _
''                )
'        End Select
'    End If
'End Function

'Private Function mGetBarcodeImageFromSheet(strValue As Variant, sht As Worksheet)
'         mGetBarcodeImageFromSheet = _
'            mGetBarcodeImageUrl( _
'                strValue, _
'                gLngBarcodeWidth, _
'                gLngBarcodeHeight, _
'                gLngBarcodeFontSize, _
'                glngStyleValue, _
'                gstrType, _
'                gLngBarcodeResolution _
'            )
'End Function

'Private Function mGetBarcodeImageUrl(strValue As Variant, _
'    Optional lngWidth As Long = 190, _
'    Optional lngHeight As Long = 45, _
'    Optional lngFontSize As Long = 1, _
'    Optional lngStyle As Long = 196, _
'    Optional strType As String = "C128B", _
'    Optional lngBarcodeResolution As Long = 1) _
'As String
'
'    If lngStyle = 0 Then
'        lngStyle = 196
'    End If
'    If Len(strValue) = 0 Then
'        mGetBarcodeImageUrl = ""
'    Else
'        mGetBarcodeImageUrl = _
'            "<img src=""" & gstrPhpBarcodeGeneratorServiceProvider & "?" & _
'                "code=" & strValue & _
'                "&style=" & lngStyle & _
'                "&type=" & strType & _
'                "&width=" & lngWidth & _
'                "&height=" & lngHeight & _
'                "&xres=" & lngBarcodeResolution & _
'                "&font=" & lngFontSize & """" & _
'            " border=""0"">"
'    End If
'End Function


