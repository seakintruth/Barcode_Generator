Attribute VB_Name = "basFileFolder"
Option Explicit
'Authored 2015-2017 by Jeremy Dean Gerdes <jeremy.gerdes@navy.mil>
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
Private Const ERROR_MORE_DATA As Long = 234
Private Const ERROR_SUCCESS As Long = 0
'[TODO] update this content to work for 64 bit Office installs

Private Declare Sub apiCopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" ( _
    ByRef lpvDest As Any, ByRef lpvSource As Any, ByVal cbCopy As Long)

Private Declare Function apiWNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionW" ( _
    ByVal lpszLocalName As Long, ByVal lpszRemoteName As Long, ByRef lngRemoteName As Long) As Long

Private Declare Function apiPathStripToRoot Lib "Shlwapi.dll" Alias "PathStripToRootW" ( _
    ByVal pPath As Long) As Long
    
Private Declare Function apiPathIsUNC Lib "Shlwapi.dll" Alias "PathIsUNCW" ( _
    ByVal pszPath As Long) As Long

Private Declare Function apiSysAllocString Lib "oleaut32.dll" Alias "SysAllocString" ( _
    ByVal pwsz As Long) As Long

Private Declare Function apiPathSkipRoot Lib "Shlwapi.dll" Alias "PathSkipRootW" ( _
    ByVal pPath As Long) As Long

Private Declare Function apiPathIsNetworkPath Lib "Shlwapi.dll" Alias "PathIsNetworkPathW" ( _
    ByVal pszPath As Long) As Long

Public Declare Function URLDownloadToFileA Lib "urlmon" (ByVal pCaller As Long, _
    ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, _
    ByVal lpfnCB As Long) _
As Long

Private Declare Function apilstrlen Lib "kernel32.dll" Alias "lstrlenW" ( _
    ByVal lpString As Long) As Long

Private Enum uriType
    uriFile = 1
    uriDirectory = 2
    uriHttp = 3
    uriUndefined = 4
End Enum

Public Const ForReading = 1
Public Const ForWriting = 2
Public Const ForAppending = 8

Public Function RemoveForbiddenFilenameCharacters(strFileName As String) As String
'https://msdn.microsoft.com/en-us/library/windows/desktop/aa365247(v=vs.85).aspx
'< (less than)
'> (greater than)
': (colon)
'" (double quote)
'/ (forward slash)
'\ (backslash)
'| (vertical bar or pipe)
'? (question mark)
'* (asterisk)
Dim strForbidden As Variant
    For Each strForbidden In Array("/", "\", "|", ":", "*", "?", "<", ">", """")
        strFileName = Replace(strFileName, strForbidden, "_")
    Next
    RemoveForbiddenFilenameCharacters = strFileName
End Function

Public Function DownloadTemplateToTemp(strPath As String, Optional strDownloadDirectory As String = "")
'Returns a directory that is the parent to the copy of the folder 'Templates'
Dim strTargetDirectoryParent As String
    If Len(strDownloadDirectory) > 0 Then 'should be validating that the dir exists and we can write to it.
        strTargetDirectoryParent = GetRelativePathViaParent(strDownloadDirectory)
    Else
        strTargetDirectoryParent = Environ$("TEMP")
    End If
    Dim strTargetDirectory As String
    strTargetDirectory = strTargetDirectoryParent & "\Templates"
    BuildDir strTargetDirectory
    If FolderExists(strTargetDirectory) Then
        CopyFolder gstrAveryTemplatePath, strTargetDirectory
    End If
    DownloadTemplateToTemp = strTargetDirectoryParent
End Function

Public Function DownloadUriFileToTemp( _
    ByVal strUrl As String, _
    Optional ByVal strDestinationExtension As String = "", _
    Optional strDownloadDirectory As String) _
As String
'Avery recently changed there download process, now relying on server files specified on sheet 'StaticValues'
    Dim lngRetVal As Long
    Dim strTempFilePath As String
    Dim strTargetDirectory As String
    strTempFilePath = Left(RemoveForbiddenFilenameCharacters(Right(strUrl, Len(strUrl) - InStrRev(strUrl, "/"))), 30)
    If Len(strDownloadDirectory) > 0 Then 'should be validating that the dir exists and we can write to it.
        strTargetDirectory = GetRelativePathViaParent(strDownloadDirectory)
    Else
        strTargetDirectory = Environ$("TEMP")
    End If
    If Len(strDestinationExtension) > 0 Then
        If Left(strDestinationExtension, 1) <> "." Then
            strDestinationExtension = "." & strDestinationExtension
        End If
    Else
    End If
    strTempFilePath = strTargetDirectory & "\" & strTempFilePath & strDestinationExtension
    'Strip HTML
    strUrl = Replace(strUrl, "<img src=""", "")
    strUrl = Replace(strUrl, """ border=""0"">", "")
    DeleteFile strTempFilePath
    lngRetVal = URLDownloadToFileA(0, strUrl, strTempFilePath, 0, 0)
    If Not FileExists(strTempFilePath) Then
        SetOriginalAppOptions
        Application.StatusBar = "File download failed. Check that you are connected to the Internet:" & strUrl
        DoEvents
        SetCustomAppOptions
    End If
    DownloadUriFileToTemp = strTempFilePath
End Function

Public Function DeleteFile(strPath As String) As Boolean
    On Error Resume Next
    Dim FSO As Object ' As Scripting.FileSystemObject
    Set FSO = CreateObject("Scripting.FileSystemObject") ' New Scripting.FileSystemObject
    If FSO.FileExists(strPath) Then
        FSO.DeleteFile strPath
    End If
    DeleteFile = Err.Number = 0
    Set FSO = Nothing
End Function

' Return true if folder exists and false if folder does not exist
Public Function FolderExists(ByVal strPath As String) As Boolean
Dim FSO As Object
    ' Note*: We used to use the vba.Dir function but using that function
    ' will lock the folder and prevent it from being deleted.
    Set FSO = CreateObject("Scripting.FileSystemObject")
    FolderExists = FSO.FolderExists(strPath)
    ' Clean up
    Set FSO = Nothing
End Function

' Return false if error occurs deleting file.
Public Function CopyFolder(ByRef strSourcePath As String, ByRef strDestinationPath As String) As Boolean
On Error Resume Next
Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Err.Clear
    FSO.CopyFolder strSourcePath, strDestinationPath, True
    CopyFolder = (Err.Number = 0)
    ' Clean up
    Set FSO = Nothing
End Function

Public Function BuildDir(strPath) As Boolean
    On Error Resume Next
    Dim FSO As Object ' As Scripting.FileSystemObject
    Set FSO = CreateObject("Scripting.FileSystemObject") ' New Scripting.FileSystemObject
    If Not FSO.FolderExists(strPath) Then
        Err.Clear
        Dim arryPaths As Variant
        Dim strBuiltPath As String, intDir As Integer, fRestore As Boolean: fRestore = False
        If Left(strPath, 2) = "\\" Then
            strPath = Right(strPath, Len(strPath) - 2)
            fRestore = True
        End If
        arryPaths = Split(strPath, "\")
        'Restore Server file path
        If fRestore Then
            arryPaths(0) = "\\" & arryPaths(0)
        End If
        For intDir = 0 To UBound(arryPaths)
            strBuiltPath = strBuiltPath & arryPaths(intDir)
            If Not FSO.FolderExists(strBuiltPath) Then
                FSO.CreateFolder strBuiltPath
            End If
            strBuiltPath = strBuiltPath & "\"
        Next
    End If
    BuildDir = (Err.Number = 0) 'True if no errors
    Set FSO = Nothing
End Function

Public Sub OpenFileWithExplorer(ByRef strFilePath As String, Optional ByRef fReadOnly As Boolean = True)
    Dim wshShell
    Set wshShell = CreateObject("WScript.Shell")
    wshShell.Exec ("Explorer.exe " & strFilePath)
    Set wshShell = Nothing

End Sub

Public Function IsArrayAllocated(ByRef avarArray As Variant) As Boolean
    On Error Resume Next
    ' Normally we only need to check LBound to determine if an array has been allocated.
    ' Some function such as Split will set LBound and UBound even if array is not allocated.
    ' See http://www.cpearson.com/excel/isarrayallocated.aspx for more details.
    IsArrayAllocated = IsArray(avarArray) And _
        Not IsError(LBound(avarArray, 1)) And _
        LBound(avarArray, 1) <= UBound(avarArray, 1)
End Function

Public Function GetRelativePathViaParent(Optional ByVal strPath As String, Optional fCreateDirectory As Boolean = True)
'Usage for up 2 dirs is GetRelativePathViaParent("..\..\Destination")
    Dim strVal As String
    If Left(strPath, 2) = "\\" Or Mid(strPath, 2, 1) = ":" Then 'we have a full path and can't use relative path
        strVal = strPath
    Else
        Dim strCurrentPath As String
        strCurrentPath = Application.ThisWorkbook.path
        Dim fIsServerPath As Boolean: fIsServerPath = False
         If Left(strCurrentPath, 2) = "\\" Then
             strCurrentPath = Right(strCurrentPath, Len(strCurrentPath) - 2)
             fIsServerPath = True
        End If
        Dim aryCurrentDirectory As Variant
        aryCurrentDirectory = Split(strCurrentPath, "\")
        Dim aryParentPath As Variant
        aryParentPath = Split(strPath, "..\")
        If fIsServerPath Then
            aryCurrentDirectory(0) = "\\" & aryCurrentDirectory(0)
        End If
        Dim intDir As Integer
        For intDir = 0 To UBound(aryCurrentDirectory) - IIf(IsArrayAllocated(aryParentPath), UBound(aryParentPath), 0)
            strVal = strVal & aryCurrentDirectory(intDir) & "\"
        Next
        strVal = StripTrailingBackSlash(strVal)
        If IsArrayAllocated(aryParentPath) Then
            strVal = strVal & "\" & aryParentPath(UBound(aryParentPath))
        End If
    End If
    If fCreateDirectory Then
        BuildDir strVal
    End If
    GetRelativePathViaParent = strVal

End Function

Public Function StripTrailingBackSlash(ByRef strPath As String)
        If Right(strPath, 1) = "\" Then
            StripTrailingBackSlash = Left(strPath, Len(strPath) - 1)
        Else
            StripTrailingBackSlash = strPath
        End If
End Function


'********************************************************************
'Next 4 functions Adapted from http://stackoverflow.com/questions/9724779/vba-identifying-whether-a-string-is-a-file-a-Directory-or-a-web-url
'posted March 2016 all Code contributions to stackoverflow after Feb 2016 are under the MIT license:https://opensource.org/licenses/MIT
'********************************************************************
'Some of the fuctions have been Adapted from: http://allenbrowne.com
'On the botom of http://allenbrowne.com/tips.html
'********************************************************************
'Permission
'You may freely use anything (code, forms, algorithms, ...) from these articles and sample databases for any purpose (personal, educational, commercial, resale, ...). All we ask is that you acknowledge this website in your code, with comments such as:
'Source: http://allenbrowne.com
'********************************************************************
Private Function mCheckPath(ByVal path) As uriType
    Dim retval
    Select Case True 'select case only tests one at a time and stops on the first True solution.
        Case HttpExists(path)
            retval = uriHttp
        Case FileExists(path)
            retval = uriFile
        Case FileExists(GetRelativePathViaParent(path))
            retval = uriFile
        Case DirectoryExists(path)
            retval = uriDirectory
        Case DirectoryExists(GetRelativePathViaParent(path))
            retval = uriDirectory
        Case Else
            retval = uriUndefined
    End Select
    mCheckPath = retval
End Function

Public Function GetPathWithoutRoot(ByRef strPath As String) As String
    apiCopyMemory ByVal VarPtr(GetPathWithoutRoot), apiSysAllocString(apiPathSkipRoot(StrPtr(strPath))), 4&
End Function

Public Function ConvertMappedDrivePathToUNCPath(ByRef strPath As String) As String

Dim strPathWithoutRoot As String
Dim strLocalPathRoot As String
Dim lngRemotePathRootLength As Long
Dim strRemotePathRoot As String
Dim lngResult As Long

    ' Convert mapped drive in path to UNC path if needed

    If apiPathIsNetworkPath(StrPtr(strPath)) = 1 Then
        If apiPathIsUNC(StrPtr(strPath)) <> 1 Then
            ' Break path into root and non-root parts
            strPathWithoutRoot = GetPathWithoutRoot(strPath)
            strLocalPathRoot = GetPathRoot(strPath)
            ' Move slash from end of local path root to beginning of path without root if needed
            If StrComp(Right$(strLocalPathRoot, 1), "\", vbBinaryCompare) = 0 Then
                strLocalPathRoot = Left$(strLocalPathRoot, Len(strLocalPathRoot) - 1)
                strPathWithoutRoot = "\" & strPathWithoutRoot
            End If
            ' Get remote name
            lngResult = apiWNetGetConnection(StrPtr(strLocalPathRoot), StrPtr(vbNullString), lngRemotePathRootLength)
            If lngResult <> ERROR_MORE_DATA Then _
                Err.Raise Number:=vbObjectError + 1, Description:="apiWNetGetConnection failed." 'DllErrorDescriptionUnexpected("apiWNetGetConnection failed.", lngResult)
            strRemotePathRoot = String$(lngRemotePathRootLength - 1, vbNullChar) ' Minus 1 because length includes terminating null character
            lngResult = apiWNetGetConnection(StrPtr(strLocalPathRoot), StrPtr(strRemotePathRoot), lngRemotePathRootLength)
            If lngResult <> ERROR_SUCCESS Then _
                Err.Raise Number:=vbObjectError + 1, Description:="apiWNetGetConnection failed." 'DllErrorDescriptionUnexpected("apiWNetGetConnection failed.", lngResult)
            ' Return path replacing mapped drive with unc path
            ConvertMappedDrivePathToUNCPath = strRemotePathRoot & strPathWithoutRoot
        Else
            ConvertMappedDrivePathToUNCPath = strPath
        End If
    Else
        ConvertMappedDrivePathToUNCPath = strPath
    End If

End Function

Public Function FileExists(ByVal strFile As String, Optional bFindDirectories As Boolean) As Boolean
    'Purpose:   Return True if the file exists, even if it is hidden.
    'Arguments: strFile: File name to look for. Current directory searched if no path included.
    '           bFindDirectories. If strFile is a Directory, FileExists() returns False unless this argument is True.
    'Note:      Does not look inside subdirectories for the file.
    'Author:    Allen Browne. http://allenbrowne.com June, 2006.
    Dim lngAttributes As Long

    'Include read-only files, hidden files, system files.
    lngAttributes = (vbReadOnly Or vbHidden Or vbSystem)
    If bFindDirectories Then
        lngAttributes = (lngAttributes Or vbDirectory) 'Include Directories as well.
    Else
        'Strip any trailing slash, so Dir does not look inside the Directory.
        Do While Right$(strFile, 1) = "\"
            strFile = Left$(strFile, Len(strFile) - 1)
        Loop
    End If
    'If Dir() returns something, the file exists.
    On Error Resume Next
    FileExists = (Len(Dir(strFile, lngAttributes)) > 0)
End Function

Public Function GetPathRoot(ByVal strPath As String) As String
   apiPathStripToRoot StrPtr(strPath)
   GetPathRoot = Left$(strPath, apilstrlen(StrPtr(strPath)))
End Function

Public Function DirectoryExists(ByVal strPath As String) As Boolean
    On Error Resume Next
    DirectoryExists = ((GetAttr(strPath) And vbDirectory) = vbDirectory)
End Function

Private Function HttpExists(ByVal sURL As String) As Boolean
    'TODO have not built out how to query that an FTP file is present, for FTP responces see: https://tools.ietf.org/html/rfc959
    On Error GoTo HandleError
    Dim oXHTTP As Object
    On Error Resume Next
    Set oXHTTP = CreateObject("MSXML2.XMLHTTP")
    If Err.Number <> 0 Then
        Set oXHTTP = CreateObject("MSXML2.SERVERXMLHTTP")
    End If
    If Not UCase(sURL) Like "HTTP:*" _
        And Not UCase(sURL) Like "HTTPS:*" _
    Then
        sURL = "https://" & sURL
    End If
    oXHTTP.Open "HEAD", sURL, False
    oXHTTP.send
    Select Case oXHTTP.Status
        Case 200 To 399, 403, 426 ' maybe 100 and 101?
            'The 2xx (Successful) class of status code indicates that the client's
            'request was successfully received, understood, and accepted. https://tools.ietf.org/html/rfc7231#section-6.3
            '403 status code indicates that the server understood the request but refuses to authorize it
            '426 tells us it's here but you need to upgrade current protocol
            HttpExists = True
        Case Else '400, 404,410 500's
            HttpExists = False
    End Select
    Exit Function
HandleError:
    Debug.Print Err.Description
    HttpExists = False
End Function



