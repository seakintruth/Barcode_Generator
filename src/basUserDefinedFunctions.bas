Attribute VB_Name = "basUserDefinedFunctions"
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
        
Public Function GetStyleOption(lngOptionValue As Long) As String
Application.Volatile (True)
    gLngOptionStyleValue = lngOptionValue
    Select Case lngOptionValue
        Case gOptStyleDisplayText
            GetStyleOption = gstrOptionStyleDisplayText
        Case gOptStyleStretchText
            GetStyleOption = gstrOptionStyleStretch
        Case gOptStyleNoText
            GetStyleOption = gstrOptionStyleNoText
    End Select
End Function


