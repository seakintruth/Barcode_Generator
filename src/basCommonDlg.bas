Attribute VB_Name = "basCommonDlg"
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
' 2011 By Bradley M. Gough modified for Excel by Jeremy Gerdes employees of the U.S. Govt
' Portions from "VBA Developer's Handbook, 2nd Edition" with permission
' by Ken Getz and Mike Gilbert
' Copyright 2001; Sybex, Inc. All rights reserved.

' Open/Save dialog flags
Private Const OFN_READONLY As Long = &H1
Private Const OFN_OVERWRITEPROMPT As Long = &H2
Private Const OFN_HIDEREADONLY As Long = &H4
Private Const OFN_NOCHANGEDIR As Long = &H8
Private Const OFN_SHOWHELP As Long = &H10
Private Const OFN_NOVALIDATE As Long = &H100
Private Const OFN_ALLOWMULTISELECT As Long = &H200
Private Const OFN_EXTENSIONDIFFERENT As Long = &H400
Private Const OFN_PATHMUSTEXIST As Long = &H800
Private Const OFN_FILEMUSTEXIST As Long = &H1000
Private Const OFN_CREATEPROMPT As Long = &H2000
Private Const OFN_SHAREAWARE As Long = &H4000
Private Const OFN_NOREADONLYRETURN As Long = &H8000
Private Const OFN_NOTESTFILECREATE As Long = &H10000
Private Const OFN_NONETWORKBUTTON As Long = &H20000
Private Const OFN_NOLONGNAMES As Long = &H40000
' Flags for hook functions and dialog templates
'Private Const OFN_ENABLEHOOK As Long = &H20
'Private Const OFN_ENABLETEMPLATE As Long = &H40
'Private Const OFN_ENABLETEMPLATEHANDLE As Long = &H80
' Windows 95 flags
Private Const OFN_EXPLORER As Long = &H80000
Private Const OFN_NODEREFERENCELINKS As Long = &H100000
Private Const OFN_LONGNAMES As Long = &H200000

' Custom flag combinations
Private Const OFN_OPENEXISTING As Long = OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY
Private Const OFN_SAVENEW As Long = OFN_PATHMUSTEXIST Or OFN_OVERWRITEPROMPT Or OFN_HIDEREADONLY
Private Const OFN_SAVENEWPATH As Long = OFN_OVERWRITEPROMPT Or OFN_HIDEREADONLY

' Message the code's reacting to.
Private Const WM_INITDIALOG As Long = &H110

Private Declare Function apiGetActiveWindow Lib "user32.dll" Alias "GetActiveWindow" () As Long

Private Declare Function apilstrlen Lib "kernel32.dll" Alias "lstrlenW" ( _
    ByVal lpString As Long) As Long

Public Function CommonDlgCallback( _
    ByVal hwnd As Long, ByVal uiMsg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long

On Error GoTo HandleError

    ' Callback function for Font and Color dialogs. The
    ' parameters for this function are defined and
    ' dictated by the API functions. You may not
    ' alter anything in the declaration besides the
    ' names, or this code will not work.

    Select Case uiMsg
        Case WM_INITDIALOG
            ' On initialization, center the dialog.
            CenterWindow hwnd
            ' You could get many other messages here, too.
            ' All the normal window messages get
            ' filtered through here, and you can
            ' react to any that you like.
    End Select

    ' Tell the original code to handle the message, too.
    ' Otherwise, things get pretty ugly.
    ' To do that, return 0.
    CommonDlgCallback = 0

ExitHere:
    Exit Function

HandleError:
    Resume ExitHere
End Function

Public Function FileDialog( _
    Optional ByRef strInitDir As String, _
    Optional ByRef strFilter As String = "All Files (*.*)" & vbNullChar & "*.*" & vbNullChar & vbNullChar, _
    Optional ByRef intFilterIndex As Integer = 1, _
    Optional ByRef strDefaultExt As String = vbNullString, _
    Optional ByRef strFileName As String = vbNullString, _
    Optional ByRef strDialogTitle As String = "Open File", _
    Optional ByRef fOpenFile As Boolean = True, _
    Optional ByRef lngFlags As Long = OFN_OPENEXISTING) As String

On Error GoTo HandleError

Dim cdl As CommonDlg

    Set cdl = New CommonDlg

    With cdl

        ' Set intial directory
        ' Note: We use current directory if argument not passed
        If strInitDir = vbNullString Then _
            strInitDir = CurDir
        .InitDir = strInitDir
        ' Set various dialog properties
        .Filter = strFilter
        .FilterIndex = intFilterIndex
        .DefaultExt = strDefaultExt
        .FileName = strFileName
        .DialogTitle = strDialogTitle
        .OpenFlags = lngFlags
        ' Enable cancel error so we can detect if cancel is selected
        .CancelError = True
        ' Set callback address
        .Callback = AddressOfToLong(AddressOf CommonDlgCallback)
        ' Set owner so dialog is modal
        .Owner = apiGetActiveWindow()

        ' Show dialog
        If fOpenFile Then
            .ShowOpen
        Else
            .ShowSave
        End If

        ' Return any flags to the calling procedure
        lngFlags = .OpenFlags

        ' Return file name
        ' Note: Trim null is only required if mulitiple files
        ' are selected but we always call it.
        FileDialog = Left$(.FileName, apilstrlen(StrPtr(.FileName)))

    End With

ExitHere:
    Set cdl = Nothing
    Exit Function

HandleError:
    Select Case Err.Number
        Case 32755 ' Cancel was selected.
            ' Cancel was selected to close dialog
            FileDialog = vbNullString
        Case Else
            ' Error handling has not been added yet
            Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End Select
    Resume ExitHere

End Function

Public Function AddressOfToLong(ByRef lngAddress As Long) As Long
    AddressOfToLong = lngAddress
End Function

Public Function ColorDialog(Optional ByRef lngColor As Long = 0) As Variant

Dim cdl As CommonDlg

On Error GoTo HandleError

    Set cdl = New CommonDlg

    With cdl

        ' Set CommonDlg color to passed color
        .Color = lngColor
        ' Enable cancel error so we can detect if cancel is selected
        .CancelError = True
        ' Ensure passed color is selected when color dialog is opened
        .ColorFlags = cdlCCRGBInit Or cdlCCEnableHook
        ' Set callback address
        .Callback = AddressOfToLong(AddressOf CommonDlgCallback)
        ' Set owner so dialog is modal
        .Owner = apiGetActiveWindow()

        ' Show dialog
        .ShowColor

        ' Return color
        ColorDialog = .Color

    End With

ExitHere:
    Set cdl = Nothing
    Exit Function

HandleError:
    Select Case Err.Number
        Case 32755 ' Cancel was selected.
            ' Cancel was selected to close dialog
            ColorDialog = Null
        Case Else
            ' Error handling has not been added yet
            Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End Select
    Resume ExitHere

End Function

Public Function FontDialog(ByRef fnt As StdFont, ByRef lngColor As Long) As Boolean

Dim cdl As CommonDlg

On Error GoTo HandleError

    Set cdl = New CommonDlg

    With cdl

        ' Set various dialog properties
        If Not fnt Is Nothing Then
            .FontBold = fnt.Bold
            .FontItalic = fnt.Italic
            .FontUnderline = fnt.Underline
            .FontStrikeThrough = fnt.Strikethrough
            .FontName = fnt.Name
            .FontScript = fnt.Weight
            .FontSize = fnt.Size
            '.FontStyle =
            .FontWeight = fnt.Weight
        End If
        ' Set font color
        .FontColor = lngColor
        ' Enable cancel error so we can detect if cancel is selected
        .CancelError = True
        ' Enable font dialog options such as font color
        .FontFlags = cdlCFEffects Or cdlCFEnableHook
        ' Set callback address
        .Callback = AddressOfToLong(AddressOf CommonDlgCallback)
        ' Set owner so dialog is modal
        .Owner = apiGetActiveWindow()

        ' Show dialog
        .ShowFont

    End With

    ' Return font properties
    With fnt
        .Bold = cdl.FontBold
        .Italic = cdl.FontItalic
        .Underline = cdl.FontUnderline
        .Strikethrough = cdl.FontStrikeThrough
        .Name = cdl.FontName
        .Weight = cdl.FontScript
        .Size = cdl.FontSize
        '.FontStyle =
        .Weight = cdl.FontWeight
    End With

    ' Return font color selected by dialog box
    lngColor = cdl.FontColor

    ' Return true to indicate cancel was not selected
    FontDialog = True

ExitHere:
    Set cdl = Nothing
    Exit Function

HandleError:
    FontDialog = False
    Select Case Err.Number
        Case 32755 ' Cancel was selected.
            ' Cancel was selected to close dialog
        Case Else
            ' Error handling has not been added yet
            Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End Select
    Resume ExitHere

End Function
