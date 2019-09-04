Attribute VB_Name = "basOpenFile"
Option Explicit
'Authored 2014-2017 Jeremy Dean Gerdes <jeremy.gerdes@navy.mil>
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

Public Sub ActionForWorkbook_BeforeSave()
On Error Resume Next
    SetEcho False
    Dim sht As Worksheet
    If Len(gstrOutputSheetName) <= 0 Then
        SetGlobalVariables
    End If
    mDeleteAllShapes ThisWorkbook.Worksheets(gstrOutputSheetName)
    Set sht = ThisWorkbook.Sheets(gstrNoticeSheetName): sht.Visible = xlSheetVisible
    sht.Activate
    Set sht = ThisWorkbook.Sheets(gstrOptionsSheetName): sht.Visible = xlSheetVeryHidden
    If Len(gstrDataSheetName) = 0 Then
        gstrDataSheetName = "Input"
    End If
    Set sht = ThisWorkbook.Sheets(gstrDataSheetName): sht.Visible = xlSheetVeryHidden
    SetEcho True
End Sub

Public Sub ActionForWorkbook_Open()
    SetCustomAppOptions
    gstrDataSheetName = "Input"
    SetGlobalVariables
    Dim sht As Worksheet
    Set sht = ThisWorkbook.Sheets(gstrDataSheetName): sht.Visible = xlSheetVisible
    sht.Activate
    Set sht = ThisWorkbook.Sheets(gstrOptionsSheetName): sht.Visible = xlSheetVeryHidden
    Set sht = ThisWorkbook.Sheets(gstrNoticeSheetName): sht.Visible = xlSheetVeryHidden
    SetOriginalAppOptions
End Sub

Public Sub OpenFileInIE(ByRef strFilePath As String)
    On Error Resume Next
    Dim objInternetExplorer 'As InternetExplorer
    'Set objInternetExplorer = New InternetExplorer
    Set objInternetExplorer = CreateObject("InternetExplorer.Application")
    objInternetExplorer.Navigate strFilePath
    objInternetExplorer.Visible = True
    Set objInternetExplorer = Nothing
    If Err.Number <> 0 Then
        OpenFileWithExplorer strFilePath
    End If
End Sub

Public Sub OpenFileInChrome(ByRef strFilePath As String)
    On Error Resume Next
    Dim strChromePath As String
    strChromePath = Environ("ProgramFiles(x86)") & "\Google\Chrome\Application\chrome.exe"
    If FileExists(strChromePath) Then
        Dim wshShell
        Set wshShell = CreateObject("WScript.Shell")
        wshShell.Exec ("""" & strChromePath & """ """ & strFilePath & """")
        Set wshShell = Nothing
    Else
        OpenFileInIE strFilePath
    End If
End Sub

Public Function GetFileExtension(ByRef strPath As String) As String
On Error Resume Next
Dim lngPosPeriod As Long
Dim lngPosSlash As Long
    ' Get start position of file extension and last slash
    lngPosPeriod = InStrRev(strPath, ".", -1, vbBinaryCompare)
    lngPosSlash = InStrRev(strPath, "\", -1, vbBinaryCompare)
    ' Verify we found a file extension
    If lngPosPeriod <> 0 And lngPosPeriod > lngPosSlash Then _
        GetFileExtension = Right$(strPath, Len(strPath) - lngPosPeriod)
End Function



