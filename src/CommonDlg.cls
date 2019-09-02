VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommonDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' 2011 By Bradley M. Gough modified for Excel by Jeremy Gerdes employees of the U.S. Govt
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
' Modified version of VBA Developer's Handbook, 2nd Edition code.

' From VBA Developer's Handbook, 2nd Edition
' by Getz and Gilbert. (Sybex)
' Copyright 2001. All rights reserved.
'
' Provide access to the File Open/Save,
' Color and Font common dialogs.
' Works similarly to the CommonDialog
' ActiveX control, but adds more features,
' and doesn't implement Printer or Help
' support.

' NOTE: This class module contains
' some redundant code (that is, code
' copied from other modules) so that
' it can be imported and used in other
' applications without needing to
' also import any subsidiary modules.

' =================
' API Constants
' =================
Private Const HWND_DESKTOP As Long = 0
Private Const LF_FACESIZE As Long = 32
Private Const FNERR_BUFFERTOOSMALL As Long = &H3003

' Modify the Open/Save dialog box.
Private Const WM_USER As Long = &H400
Private Const CDM_FIRST As Long = (WM_USER + 100)

' =================
' API Enums (values defined by API,
' Enums defined here). These are set
' up to match the CommonDialog ActiveX
' control's constants, but we've added
' some extras.
' =================

Public Enum dhCDFontType
    RASTER_FONTTYPE = &H1
    DEVICE_FONTTYPE = &H2
    TRUETYPE_FONTTYPE = &H4
    BOLD_FONTTYPE = &H100
    ITALIC_FONTTYPE = &H200
    REGULAR_FONTTYPE = &H400
    SCREEN_FONTTYPE = &H2000
    PRINTER_FONTTYPE = &H4000
    SIMULATED_FONTTYPE = &H8000
    OPENTYPE_FONTTYPE = &H10000
    TYPE1_FONTTYPE = &H20000
    DSIG_FONTTYPE = &H40000
End Enum

Public Enum dhFontFaceAPI
    ANSI_CHARSET = 0
    DEFAULT_CHARSET = 1
    SYMBOL_CHARSET = 2
    SHIFTJIS_CHARSET = 128
    HANGEUL_CHARSET = 129
    GB2312_CHARSET = 134
    CHINESEBIG5_CHARSET = 136
    OEM_CHARSET = 255
    JOHAB_CHARSET = 130
    HEBREW_CHARSET = 177
    ARABIC_CHARSET = 178
    GREEK_CHARSET = 161
    TURKISH_CHARSET = 162
    VIETNAMESE_CHARSET = 163
    THAI_CHARSET = 222
    EASTEUROPE_CHARSET = 238
    RUSSIAN_CHARSET = 204
    MAC_CHARSET = 77
    BALTIC_CHARSET = 186
End Enum

Public Enum dhColorConstants
    cdlCCFullOpen = 2
    cdlCCHelpButton = 8
    cdlCCPreventFullOpen = 4
    cdlCCRGBInit = 1
    cdlCCAnyColor = &H100
    cdlCCEnableHook = &H10
    cdlCCSolidColor = &H80
End Enum

Public Enum dhErrorConstants
    cdlAlloc = 32752
    cdlBufferTooSmall = 20476
    cdlCancel = 32755
    cdlCreateICFailure = 28661
    cdlDialogFailure = -32768
    cdlDndmMismatch = 28662
    cdlFindResFailure = 32761
    cdlGetDevModeFail = 28666
    cdlGetNotSupported = 394
    cdlHelp = 32751
    cdlInitFailure = 28665
    cdlInitialization = 32765
    cdlInvalidFileName = 20477
    cdlInvalidPropertyValue = 380
    cdlInvalidSafeModeProcCall = 680
    cdlLoadDrvFailure = 28667
    cdlLoadResFailure = 32760
    cdlLoadStrFailure = 32762
    cdlLockResFailure = 32759
    cdlMemAllocFailure = 32758
    cdlMemLockFailure = 32757
    cdlNoDefaultPrn = 28663
    cdlNoDevices = 28664
    cdlNoFonts = 24574
    cdlNoInstance = 32763
    cdlNoTemplate = 32764
    cdlParseFailure = 28669
    cdlPrinterCodes = 28671
    cdlPrinterNotFound = 28660
    cdlRetDefFailure = 28668
    cdlSetNotSupported = 383
    cdlSetupFailure = 28670
    cdlSubclassFailure = 20478
End Enum

Public Enum dhFileOpenConstants
    cdlOFNAllowMultiselect = 512
    cdlOFNCreatePrompt = 8192
    cdlOFNEnableHook = 32
    cdlOFNEnableSizing = 8388608
    cdlOFNExplorer = 524288
    cdlOFNExtensionDifferent = 1024
    cdlOFNFileMustExist = 4096
    cdlOFNHelpButton = 16
    cdlOFNHideReadOnly = 4
    cdlOFNLongNames = 2097152
    cdlOFNNoChangeDir = 8
    cdlOFNNoDereferenceLinks = 1048576
    cdlOFNNoLongNames = 262144
    cdlOFNNoNetworkButton = 131072
    cdlOFNNoReadOnlyReturn = 32768
    cdlOFNNoValidate = 256
    cdlOFNOverwritePrompt = 2
    cdlOFNPathMustExist = 2048
    cdlOFNReadOnly = 1
    cdlOFNShareAware = 16384
End Enum

Public Enum dhFileOpenConstantsEx
    cdlOFNExNoPlacesBar = 128
End Enum

Public Enum dhFontsConstants
    cdlCFANSIOnly = &H400
    cdlCFApply = &H200
    cdlCFBoth = &H3
    cdlCFEffects = &H100
    cdlCFEnableHook = &H8
    cdlCFFixedPitchOnly = &H4000
    cdlCFForceFontExist = &H10000
    cdlCFInitToLogFontStruct = &H40
    cdlCFLimitSize = &H2000
    cdlCFNoFaceSel = &H80000
    cdlCFNoSimulations = &H1000
    cdlCFNoSizeSel = &H200000
    cdlCFNoStyleSel = &H100000
    cdlCFNoVectorFonts = &H800
    cdlCFNoVertFonts = &H1000000
    cdlCFPrinterFonts = &H2
    cdlCFScalableOnly = &H20000
    cdlCFScreenFonts = &H1
    cdlCFShowHelp = &H4
    cdlCFTTOnly = &H40000
    cdlCFUseStyle = &H80
    cdlCFWYSIWYG = &H8000       ' must also have cdlCFScreenFonts & cdlCFPrinterFonts
End Enum

' You can use these values in the
' File Open/Save callback function
' to modify the text or visibility
' of any of the controls on the
' dialog. See the example callback
' function for a demo.
Public Enum dhFileOpenSaveControls
    fosCurrentFolder = &H471
    fosCurrentFolderLabel = &H443
    fosContentsList = &H460
    fosContentsListLabel = &H440
    fosSelectedFile = &H480
    fosSelectedFileLabel = &H442
    fosFilterList = &H470
    fosFilterListLabel = &H441
    fosReadOnly = &H410
    fosOKButton = 1
    fosCancelButton = 2
    fosHelpButton = &H40E
End Enum

Public Enum dhCommonDialogManage
    CDM_SETCONTROLTEXT = (CDM_FIRST + &H4)
    CDM_HIDECONTROL = (CDM_FIRST + &H5)
End Enum

' =================
' API Types
' =================
Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(1 To LF_FACESIZE) As Byte
End Type

Private Type ChooseColor
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As Long
    flags As dhColorConstants
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As dhFileOpenConstants
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type OPENFILENAMEEX
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As dhFileOpenConstants
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
    ' The following extra elements are valid
    ' only if running under Windows 2000/XP
    pvReserved As Long
    dwReserved As Long
    FlagsEx As Long
End Type

Private Type ChooseFont
    lStructSize As Long
    hWndOwner As Long            '  caller's window handle
    hdc As Long                  '  printer DC/IC or NULL
    lpLogFont As Long
    iPointSize As Long           '  10 * size in points of selected font
    flags As dhFontsConstants    '  enum. type flags
    rgbColors As Long            '  returned text color
    lCustData As Long            '  data passed to hook fn.
    lpfnHook As Long             '  ptr. to hook function
    lpTemplateName As String     '  custom template name
    hInstance As Long            '  instance handle of.EXE that
                                 '  contains cust. dlg. template
    lpszStyle As String          '  return the style field here
                                 '  must be LF_FACESIZE or bigger
    nFontType As Integer         '  same value reported to the EnumFonts
                                 '  call back with the extra FONTTYPE_
                                 '  bits added
    MISSING_ALIGNMENT As Integer
    nSizeMin As Long             '  minimum pt size allowed &
    nSizeMax As Long             '  max pt size allowed if
                                 '  CF_LIMITSIZE is used
End Type

Private mblnIsWin2000 As Boolean

' =================
' API Declarations
' =================
Private Declare Function GetDC Lib "user32.dll" ( _
    ByVal hwnd As Long) As Long

Private Declare Function ReleaseDC Lib "user32.dll" ( _
    ByVal hwnd As Long, ByVal hdc As Long) As Long

Private Const LOGPIXELSY As Long = 90

Private Declare Function GetDeviceCaps Lib "gdi32" ( _
    ByVal hdc As Long, ByVal nIndex As Long) As Long

Private Declare Function MulDiv Lib "kernel32.dll" ( _
    ByVal nNumber As Long, ByVal nNumerator As Long, _
    ByVal nDenominator As Long) As Long

Private Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long

Private Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" ( _
    pChoosefont As ChooseFont) As Long

Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" ( _
    pChoosecolor As ChooseColor) As Long

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" ( _
    pOpenfilename As Any) As Long

Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" ( _
    pOpenfilename As Any) As Long

' =================
' Storage for property values.
' =================

' Returns/sets the size of the file name
' buffer to use for the FileOpen dialog box.
' The default size is 2000.
Public FileNameBufferSize As Long

' Returns/sets the custom file open/save filter.
' Public CustomFilter As String

' Returns/sets the default filename extension for the dialog box.
Public DefaultExt As String

' Sets the string displayed in the title bar of the dialog box.
Public DialogTitle As String

' Returns/sets the path and filename of a selected file.
Public FileName As String

' Returns/sets the name (without the path) of the file to open or save at run time.
Public FileTitle As String

' Returns/sets the filters that are displayed in the Type list box of a dialog box.
Public Filter As String

' Returns/sets a default filter for an Open or Save As dialog box.
Public FilterIndex As Long

' Returns/sets the initial file directory.
Public InitDir As String

' Returns/sets the selected color.
Public Color As Long

' Sets the hWnd of the dialog owner.
Public hWndOwner As Long

' Sets/Returns the character set.
' Although interesting, doesn't correspond
' to any property in the host app.
Public FontScript As dhFontFaceAPI

' Text describing the selected font style.
Public FontStyle As String

' Set/Returns the minimum and maximum font sizes,
' if you've set the cdlCFLimitSize flag.
' Disregarded otherwise.
Public Min As Integer
Public Max As Integer

' Returns the selected font color.
Public FontColor As Long

' Flag settings (for backwards compatability only)
Public flags As Long

' Flags specific to the specific dialog box.
Public FontFlags As dhFontsConstants
Public ColorFlags As dhColorConstants
Public OpenFlags As dhFileOpenConstants
Public OpenFlagsEx As dhFileOpenConstantsEx

' Address of the callback function.
Public Callback As Long

' Specifies the name of the font that appears in each row for the given level.
Public FontName As String

' Indicates whether an error is generated when the user chooses the Cancel button.
Public CancelError As Boolean

' Returns/sets italic font styles.
Public FontItalic As Boolean

' Returns/sets bold font styles. Included for
' backwards compatability. Use FontWeight
' instead.
Public FontBold As Boolean

' Font weight, from 100 to 900 (in multiples of 100)
' 700 is bold, 400 is normal.
Public FontWeight As Long

' Specifies the size (in points) of the font that appears in each row for the given level.
Public FontSize As Single

' Returns/sets strikethrough font styles.
Public FontStrikeThrough As Boolean

' Returns/sets underline font styles.
Public FontUnderline As Boolean

' Retrieve the font type, from the dhCDFontType
' list of options. Can be any number of
' items from the group, OR'd together.
Private mlngFontType As dhCDFontType

' Retrieve the 16 user-defined colors
' returned from the color chooser dialog.
Private malngColors(0 To 15) As Long

' Retrieve the offset within the full file name
' to the file portion, or the extension portion.
Private mlngFileOffset As Long
Private mlngFileExtOffset As Long

' Retrieve the list of files selected
' if cdlOFNAllowMultiSelect flag
' is set. If not, this array contains
' only the path, and single file selected.
Private mastrFileList() As String

Private Sub Class_Initialize()

    ' Note: If Class_Initialize errors are NOT handled by this procedure they will automatically
    ' be handled by the procedure that creates a class instance or by the procedure that first
    ' uses class instance if class declared with module or global scope using 'New' keyword.

    #If DebugEnabled Then
        'DebugClassInitialize "CommonDlg"
    #End If

    ' Assume the default size.
    FileNameBufferSize = 20000

End Sub

Private Sub Class_Terminate()

On Error GoTo HandleError

    ' Note: Class_Terminate errors are NOT handled by the procedure that destroyed class instance
    ' because class instances can automatically be destroyed by the garbage collector.

    #If DebugEnabled Then
        'DebugClassTerminate "CommonDlg"
    #End If

ExitHere:
    Exit Sub

HandleError:
   

End Sub

' bmg@2007
Public Property Get Owner() As Long
    Owner = hWndOwner
End Property

' bmg@2007
Public Property Let Owner(ByRef hwnd As Long)
    hWndOwner = hwnd
End Property

Public Property Get FileList() As String()
    ' Get the parsed list of files.
    ' If there are items in this list,
    ' the 0th element is the path, and the
    ' rest are the selected files.
    ' Even if you only select a single
    ' file, we populate this array.
    FileList = mastrFileList
End Property

Public Property Get FileOffset() As Long
    ' Returns the offset within the full file name
    ' to the file portion.
    FileOffset = mlngFileOffset
End Property

Public Property Get FileExtOffset() As Long
    ' Returns the offset within the full file name
    ' to the file portion.
    FileExtOffset = mlngFileExtOffset
End Property

Public Property Get CustomColors() As Long()
    ' Return the array of custom colors.
    CustomColors = malngColors
End Property

Public Property Let CustomColors(ByRef alngColors() As Long)
    Dim i As Integer

    ' The array passed in must be indexed from
    ' 0 to 15. If not, weird things are going
    ' to happen -- we just copy from those
    ' indexes directly over.
    On Error GoTo HandleError
    For i = 0 To 15
        malngColors(i) = alngColors(i)
NextValue:
    Next i

ExitHere:
    Exit Property

HandleError:
    Resume NextValue
End Property

Public Property Get FontType() As dhCDFontType
    FontType = mlngFontType
End Property

' =================
' CommonDlg Methods
' =================
Public Sub ShowColor()

    ' From VBA Developer's Handbook, 2nd Edition
    ' by Getz and Gilbert. (Sybex)
    ' Copyright 2001. All rights reserved.
    '
    ' Displays the CommonDialog control's Color dialog box.

    Dim cc As ChooseColor

    SetColorProperties cc
    If ChooseColor(cc) <> 0 Then
        GetColorProperties cc
    Else
        ' If the user wants to raise an error for the Escape
        ' do it now.
        If CancelError Then
            Err.Raise cdlCancel, , "Cancel was selected."
        End If
    End If

End Sub

Public Sub ShowFont()

    ' From VBA Developer's Handbook, 2nd Edition
    ' by Getz and Gilbert. (Sybex)
    ' Copyright 2001. All rights reserved.
    '
    ' Display the CommonDialog control's Font dialog box

    Dim cf As ChooseFont
    Dim lf As LOGFONT
    Dim strStyle As String

    ' Arbitrarily allow 100 characters
    ' for the style string.
    strStyle = Space$(100)
    SetFontProperties cf, lf, strStyle
    If ChooseFont(cf) <> 0 Then
        ' The user pressed the OK button
        GetFontProperties cf, lf
    Else
        ' If the user wants to raise an error for the Escape
        ' do it now.
        If CancelError Then
            Err.Raise cdlCancel, , "Cancel was selected."
        End If
    End If
End Sub

Public Sub ShowOpen()

    ' From VBA Developer's Handbook, 2nd Edition
    ' by Getz and Gilbert. (Sybex)
    ' Copyright 2001. All rights reserved.
    '
    ' Displays the CommonDialog control's Open dialog box.

    Dim ofn As OPENFILENAMEEX
    Dim lngErr As Long

    SetOpenProperties ofn
    If GetOpenFileName(ofn) <> 0 Then
        GetOpenProperties ofn
    Else
        lngErr = CommDlgExtendedError()
        Select Case lngErr
            Case FNERR_BUFFERTOOSMALL
                Err.Raise cdlBufferTooSmall, , "File name buffer is too small for the selected files."
            Case 0
                ' If the user wants to raise an error for the Escape
                ' do it now.
                If CancelError Then
                    Err.Raise cdlCancel, , "Cancel was selected."
                End If
            Case Else
                Err.Raise lngErr, , "Unexpected error."
        End Select
    End If
End Sub

Public Sub ShowSave()

    ' From VBA Developer's Handbook, 2nd Edition
    ' by Getz and Gilbert. (Sybex)
    ' Copyright 2001. All rights reserved.
    '
    ' Displays the CommonDialog control's Save As dialog box.

    Dim ofn As OPENFILENAMEEX
    Dim lngErr As Long

    SetOpenProperties ofn
    If GetSaveFileName(ofn) <> 0 Then
        GetOpenProperties ofn
    Else
        lngErr = CommDlgExtendedError()
        Select Case lngErr
            Case FNERR_BUFFERTOOSMALL
                Err.Raise cdlBufferTooSmall, , "File name buffer is too small for the selected files."
            Case 0
                ' If the user wants to raise an error for the Escape
                ' do it now.
                If CancelError Then
                    Err.Raise cdlCancel, , "Cancel was selected."
                End If
            Case Else
                Err.Raise lngErr, , "Unexpected error."
        End Select
    End If
End Sub

Private Sub SetOpenProperties(ByRef ofn As OPENFILENAMEEX)

    ' From VBA Developer's Handbook, 2nd Edition
    ' by Getz and Gilbert. (Sybex)
    ' Copyright 2001. All rights reserved.
    '

    ' Copy object properties into the data
    ' structure before calling the API.

    Dim strFileName As String
    Dim strFileTitle As String
    Dim ofnTmp As OPENFILENAME

    ' Show the Open common dialog.
    ' Allocate string space for the returned strings.
    strFileName = String$(FileNameBufferSize, vbNullChar)
    LSet strFileName = FileName & vbNullChar
    strFileTitle = String$(1024, vbNullChar)

    With ofn
        ' If we're not running Windows 2000/XP set structure
        ' size equal to the older UDT
        If mblnIsWin2000 Then
            .lStructSize = Len(ofn)
        Else
            .lStructSize = Len(ofnTmp)
        End If

        .hWndOwner = hWndOwner
        ' The API doesn't want those "|" things, it wants
        ' vbNullChar, with an extra one on the end.
        .lpstrFilter = Replace$(Trim$(Filter), "|", vbNullChar, 1, -1, vbBinaryCompare) & vbNullChar
        .nFilterIndex = FilterIndex

        .lpstrFile = strFileName
        .nMaxFile = Len(strFileName)
        .lpstrFileTitle = strFileTitle
        .nMaxFileTitle = Len(strFileTitle)
        .lpstrTitle = DialogTitle

        ' You can set either the OpenFlags
        ' or general Flags properties. We'll
        ' OR them together. If you use both, you'd
        ' better know what you're doing!
        ' In addition, we're going to assume that you
        ' always want the explorer-style interface.
        ' Can't imagine why you wouldn't, at this point.
        .flags = OpenFlags Or flags Or cdlOFNExplorer
        .FlagsEx = OpenFlagsEx
        .lpstrDefExt = DefaultExt
        .lpstrInitialDir = InitDir

        ' We don't support the CustomFilter
        ' property, but you could add it in
        ' if you like. This buffer
        ' must contain at least 40 characters
        ' to make WinNT happy.
        .lpstrCustomFilter = String$(40, vbNullChar)
        .nMaxCustFilter = Len(.lpstrCustomFilter)

        If .flags And cdlOFNEnableHook Then
            .lpfnHook = Callback
        End If
    End With
End Sub

Private Sub GetOpenProperties(ByRef ofn As OPENFILENAMEEX)

    ' From VBA Developer's Handbook, 2nd Edition
    ' by Getz and Gilbert. (Sybex)
    ' Copyright 2001. All rights reserved.
    '

    ' Retrieve properties from the API structure
    ' back into properties of this object.

    Dim astrFileInfo() As String
    Dim intPos As Integer
    Dim strFileName As String

    With ofn
        FileName = .lpstrFile
        OpenFlags = .flags
        OpenFlagsEx = .FlagsEx
        flags = .flags
        FileTitle = .lpstrFileTitle
        FilterIndex = .nFilterIndex
        mlngFileExtOffset = .nFileExtension
        mlngFileOffset = .nFileOffset
        ' CustomFilter = .lpstrCustomFilter
        If .nFileOffset > 0 Then
            strFileName = .lpstrFile
            If Mid$(strFileName, mlngFileOffset, 1) = vbNullChar Then
                ' Look for trailing double null chars, and trim
                ' the string there.
                intPos = InStr(1, strFileName, vbNullChar & vbNullChar, vbBinaryCompare)
                If intPos > 0 Then
                    strFileName = Left$(strFileName, intPos - 1)
                End If
                astrFileInfo = Split(strFileName, vbNullChar)
                mastrFileList = astrFileInfo
            Else
                ' Only a single file selected,
                ' so break it up into path and file
                ' portion, as if the user had selected
                ' multiple files.
                ReDim mastrFileList(0 To 1)
                mastrFileList(0) = Left$(strFileName, mlngFileOffset - 1)
                mastrFileList(1) = dhTrimNull(Mid$(strFileName, mlngFileOffset + 1))
                FileName = dhTrimNull(FileName)
            End If
        End If
    End With
End Sub

Private Sub SetColorProperties(ByRef cc As ChooseColor)

    ' From VBA Developer's Handbook, 2nd Edition
    ' by Getz and Gilbert. (Sybex)
    ' Copyright 2001. All rights reserved.
    '

    ' Copy object properties into the data
    ' structure before calling the API.

    cc.lStructSize = LenB(cc)
    cc.hWndOwner = hWndOwner
    cc.rgbResult = Color
    cc.lpCustColors = VarPtr(malngColors(0))

    ' You can set either the ColorFlags
    ' or general Flags properties. We'll
    ' OR them together. If you use both, you'd
    ' better know what you're doing!
    cc.flags = ColorFlags Or flags

    ' This had better be the address of
    ' a public function in a standard
    ' module, or you're going down!
    ' Use the dhFnPtrToLong procedure
    ' to convert from AddressOf to
    ' long.
    If cc.flags And cdlCCEnableHook Then
        cc.lpfnHook = Callback
    End If
End Sub

Private Sub GetColorProperties(ByRef cc As ChooseColor)
    ' From VBA Developer's Handbook, 2nd Edition
    ' by Getz and Gilbert. (Sybex)
    ' Copyright 2001. All rights reserved.
    '

    ' Retrieve properties from the API structure
    ' back into properties of this object.

    Color = cc.rgbResult
End Sub

Private Sub SetFontProperties( _
    ByRef cf As ChooseFont, ByRef lf As LOGFONT, ByRef strStyle As String)
    ' From VBA Developer's Handbook, 2nd Edition
    ' by Getz and Gilbert. (Sybex)
    ' Copyright 2001. All rights reserved.
    '

    ' Copy object properties into the data
    ' structure before calling the API.

    On Error Resume Next
    Dim lngFlags As Long

    cf.lStructSize = LenB(cf)
    If LenB(FontName) <> 0 Then
        dhSetFaceName lf, FontName
    End If
    cf.lpLogFont = VarPtr(lf)
    cf.hWndOwner = hWndOwner
    cf.lpszStyle = FontStyle

    lf.lfHeight = CalcHeightFromPoints()
    lf.lfStrikeOut = FontStrikeThrough
    lf.lfUnderline = FontUnderline
    lf.lfItalic = FontItalic
    lf.lfCharSet = FontScript

    If FontWeight = 0 Then
        If FontBold Then
            lf.lfWeight = 700
        Else
            lf.lfWeight = 400
        End If
    Else
        lf.lfWeight = FontWeight
    End If

    cf.rgbColors = FontColor
    cf.nSizeMax = Max
    cf.nSizeMin = Min

    ' You can set either the FontFlags
    ' or general Flags properties. We'll
    ' OR them together. If you use both, you'd
    ' better know what you're doing!
    ' We also OR in cdlCFInitToLogFontStruct,
    ' 'cause you generally want to do that.

    ' In addition, if the user hasn't specified
    ' either/both cdlCFPrinterFonts or cdlCFScreenFonts
    ' we're going to assume they want both.
    lngFlags = flags Or FontFlags
    If Not (lngFlags And cdlCFPrinterFonts) And _
        Not (lngFlags And cdlCFScreenFonts) Then
        lngFlags = lngFlags Or cdlCFBoth
    End If
    cf.flags = lngFlags Or cdlCFInitToLogFontStruct

    ' This had better be the address of
    ' a public function in a standard
    ' module, or you're going down!
    ' Use the dhFnPtrToLong procedure
    ' to convert from AddressOf to
    ' long.
    If cf.flags And cdlCFEnableHook Then
        cf.lpfnHook = Callback
    End If
End Sub

Private Sub GetFontProperties(ByRef cf As ChooseFont, ByRef lf As LOGFONT)

    ' From VBA Developer's Handbook, 2nd Edition
    ' by Getz and Gilbert. (Sybex)
    ' Copyright 2001. All rights reserved.
    '
    ' Retrieve properties from the API structure
    ' back into properties of this object.

    On Error Resume Next
    FontName = dhTrimNull(StrConv(lf.lfFaceName, vbUnicode))
    FontColor = cf.rgbColors
    FontItalic = lf.lfItalic
    FontBold = ((cf.nFontType And BOLD_FONTTYPE) <> 0)
    FontWeight = lf.lfWeight
    FontSize = cf.iPointSize \ 10
    FontStrikeThrough = lf.lfStrikeOut
    FontUnderline = lf.lfUnderline
    FontScript = lf.lfCharSet
    FontStyle = dhTrimNull(cf.lpszStyle)
    mlngFontType = cf.nFontType
End Sub

Private Function CalcHeightFromPoints() As Long

    ' From VBA Developer's Handbook, 2nd Edition
    ' by Getz and Gilbert. (Sybex)
    ' Copyright 2001. All rights reserved.
    '
    Dim hdc As Long
    Dim lngLogPixelsY As Long

    On Error GoTo HandleError

    ' Assume an invalid value for failure.
    CalcHeightFromPoints = 0

    ' Convert from points back to the internal
    ' device units value.
    hdc = GetDC(HWND_DESKTOP)
    If hdc <> 0 Then
        lngLogPixelsY = GetDeviceCaps(hdc, LOGPIXELSY)
        CalcHeightFromPoints = _
            -1 * MulDiv(CInt(FontSize), lngLogPixelsY, 72)
    End If

ExitHere:
    Exit Function

HandleError:
    Resume ExitHere
End Function

Private Function dhTrimNull(ByRef strVal As String) As String
    ' From VBA Developer's Handbook, 2nd Edition
    ' by Getz and Gilbert. (Sybex)
    ' Copyright 2001. All rights reserved.

    ' Trim the end of a string, stopping at the first
    ' null character.

    Dim intPos As Integer
    intPos = InStr(1, strVal, vbNullChar, vbBinaryCompare)
    Select Case intPos
        Case Is > 1
            dhTrimNull = Left$(strVal, intPos - 1)
        Case 0
            dhTrimNull = strVal
        Case 1
            dhTrimNull = vbNullString
    End Select
End Function

Private Sub dhSetFaceName(ByRef lf As LOGFONT, ByRef strValue As String)

    ' From VBA Developer's Handbook, 2nd Edition
    ' by Getz and Gilbert. (Sybex)
    ' Copyright 2001. All rights reserved.
    '
    ' Given a string, get it back into the ANSI byte array
    ' contained within a LOGFONT structure.

    Dim intLen As Integer
    Dim intI As Integer
    Dim abytTemp() As Byte

    On Error GoTo HandleError

    abytTemp = StrConv(strValue, vbFromUnicode)
    intLen = UBound(abytTemp) + 1

    ' Make sure the string isn't too long.
    If intLen > LF_FACESIZE - 1 Then
        intLen = LF_FACESIZE - 1
    End If
    For intI = 1 To intLen
        lf.lfFaceName(intI) = abytTemp(intI - 1)
    Next intI
    ' Tack on a final Chr$(0).
    lf.lfFaceName(intI) = 0

ExitHere:
    Exit Sub

HandleError:
    Resume ExitHere
End Sub

