Attribute VB_Name = "basCenterWindow"
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

Private Type POINT
    X As Long
    Y As Long
End Type

' Window Style constants
Private Const WS_CHILD As Long = &H40000000

Private Declare Function apiGetWindowLong Lib "user32.dll" Alias "GetWindowLongW" ( _
    ByVal hwnd As Long, ByVal nIndex As Long) As Long

Private Declare Function apiGetWindowRect Lib "user32.dll" Alias "GetWindowRect" ( _
    ByVal hwnd As Long, ByRef lpRect As RECT) As Long

Private Declare Function apiGetParent Lib "user32.dll" Alias "GetParent" ( _
    ByVal hwnd As Long) As Long

Private Declare Function apiGetClientRect Lib "user32.dll" Alias "GetClientRect" ( _
    ByVal hwnd As Long, ByRef lpRect As RECT) As Long

' MonitorFromWindow dwFlags constants
Private Const MONITOR_DEFAULTTONEAREST As Long = &H2

Private Declare Function apiMonitorFromWindow Lib "user32.dll" Alias "MonitorFromWindow" ( _
    ByVal hwnd As Long, ByVal dwFlags As Long) As Long

Private Declare Function apiGetMonitorInfo Lib "user32.dll" Alias "GetMonitorInfoW" ( _
    ByVal hMonitor As Long, ByRef lpmi As MONITORINFO) As Long

Private Declare Function apiMoveWindow Lib "user32.dll" Alias "MoveWindow" ( _
    ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, _
    ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Private Declare Function apiGetProfileString Lib "kernel32.dll" Alias "GetProfileStringW" ( _
    ByVal lpAppName As Long, ByVal lpKeyName As Long, _
    ByVal lpDefault As Long, ByVal lpReturnedString As Long, _
    ByVal nSize As Long) As Long

' GetWindowLong nIndex constants
Private Const GWL_STYLE As Long = (-16)

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type MONITORINFO
    cbSize As Long
    rcMonitor As RECT
    rcWork As RECT
    dwFlags As Long
End Type

' This procedure is used to center and size all forms.  We often dynamically
' change the size of the form during it's load event.  This breaks the built
' in AutoCenter feature so we use this procedure.  Note that I've choosen to
' always center a form during it's load event.
' This procedure is also used to set the form size.  The form's size is always
' made a small as possible if it's a subform.  If not a subform, the form is
' given a standard boarder width (around the outer most controls).
Public Sub CenterWindow(ByRef hwnd As Long)

Dim rc As RECT
Dim lngWidth As Long
Dim lngHeight As Long
Dim hwndParent As Long
Dim rcParentClient As RECT
Dim hMonitor As Long
Dim mi As MONITORINFO
Dim lngMoveX As Long
Dim lngMoveY As Long

    ' If window is a child window, center within parent window client area
    ' If window is NOT a child window, center within desktop work area
    If apiGetWindowLong(hwnd, GWL_STYLE) And WS_CHILD Then

        ' Get window coordinates
        apiGetWindowRect hwnd, rc
        With rc
            ' Calculate window width and height
            lngWidth = (rc.Right - rc.Left)
            lngHeight = (rc.Bottom - rc.Top)
        End With

        ' Get parent window handle
        hwndParent = apiGetParent(hwnd)
        ' Get parent window client area coordinates
        apiGetClientRect hwndParent, rcParentClient

        ' Calculate X and Y move needed to move window to the
        ' center of the parent window client area
        lngMoveX = ((rcParentClient.Right - rcParentClient.Left) - lngWidth) / 2
        lngMoveY = ((rcParentClient.Bottom - rcParentClient.Top) - lngHeight) / 2

    Else

        ' Get window coordinates
        apiGetWindowRect hwnd, rc
        With rc
            ' Calculate window width and height
            lngWidth = (rc.Right - rc.Left)
            lngHeight = (rc.Bottom - rc.Top)
        End With

        ' Get parent window handle
        hwndParent = apiGetParent(hwnd)
        If hwndParent <> 0 Then
            ' Get handle to monitor parent window is open in
            hMonitor = apiMonitorFromWindow(hwndParent, MONITOR_DEFAULTTONEAREST)
        Else
            ' Get handle to monitor window is open in
            hMonitor = apiMonitorFromWindow(hwnd, MONITOR_DEFAULTTONEAREST)
        End If
        ' Initialize the MONITORINFO structure
        mi.cbSize = Len(mi)
        ' Get the monitor information
        apiGetMonitorInfo hMonitor, mi

        ' Calculate X and Y move needed to move window to the
        ' center of the work area
        lngMoveX = mi.rcWork.Left + (((mi.rcWork.Right - mi.rcWork.Left) - lngWidth) / 2)
        lngMoveY = mi.rcWork.Top + (((mi.rcWork.Bottom - mi.rcWork.Top) - lngHeight) / 2)
    End If

    ' Move window to center position, without changing z-order.
    apiMoveWindow hwnd, lngMoveX, lngMoveY, lngWidth, lngHeight, 1

End Sub

Public Function GetDefaultPrinterName() As String

' This function is from Peter Walker.
' Check out his web site at:
' http://www.users.bigpond.com/papwalker/
Dim strWindowsDevice As String
Dim lngWindowsDeviceLength As Long

    ' Call the API passing null as the parameter for the lpKeyName parameter.
    ' This causes the API to return a list of all keys under that section.

    ' Name limit for printers is 221 Characters, Add 1 for ','
    strWindowsDevice = String$(222, vbNullChar)
    lngWindowsDeviceLength = apiGetProfileString(StrPtr("Windows"), StrPtr("Device"), StrPtr(vbNullString), StrPtr(strWindowsDevice), Len(strWindowsDevice))
    If lngWindowsDeviceLength = 0 Then _
        Err.Raise Number:=vbObjectError + 1, Description:="Unable to get default printer name."
    GetDefaultPrinterName = Left$(strWindowsDevice, InStr(1, strWindowsDevice, ",", vbBinaryCompare) - 1)

End Function
