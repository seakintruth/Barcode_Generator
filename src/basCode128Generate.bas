Attribute VB_Name = "basCode128Generate"
Option Explicit
'=============================================================================================
'From https://www.mrexcel.com/forum/excel-questions/970865-excel-vba-barcode-generator.html#post4659015
'=============================================================================================
Type BarParams
    Pos As Long
    Width As Byte
End Type

Public Sub GenerateBarcodesFromInput() 'Worksheet Input
    SetEcho False
    SetGlobalVariables
    Dim shtData As Worksheet: Set shtData = ThisWorkbook.Worksheets(gstrDataSheetName)
    Dim shtOutput As Worksheet: Set shtOutput = ThisWorkbook.Worksheets(gstrOutputSheetName)
    ' Clean up the output
    mDeleteAllShapes shtOutput
    Dim rngData As Range
    Set rngData = Intersect( _
        GetNonBlankCellsFromWorksheet(shtData), _
        shtData.Range("A:A") _
    )
    If IsNothing(rngData) Then
        MsgBox "You must enter data in column A to generate barcodes", _
            vbInformation, "Barcode Generator, Need User Input"
    Else
        Dim shp As Shape
        Dim cl As Range
        Dim shpText  As Shape
        Dim aryShapeNames As Variant
        ReDim aryShapeNames(1 To 2)
        '-------------------------------------------------------------------
        ' Define the padding values and other selected values
        Dim sglLeftPadding As Single: sglLeftPadding = shtData.Range(gstrLeftPaddingImageCellName).Value
        Dim sglRightPadding As Single: sglRightPadding = shtData.Range(gstrRightPaddingImageCellName).Value
        Dim sglTopPadding As Single: sglTopPadding = shtData.Range(gstrTopPaddingImageCellName).Value
        Dim sglBottomPadding As Single: sglBottomPadding = shtData.Range(gstrBottomPaddingImageCellName).Value
        Dim sglFontSize As Single: sglFontSize = shtData.Range(gstrFontSizeCellName).Value
        Dim sglBarcodeHeight As Single: sglBarcodeHeight = shtData.Range(gstrHeightCellName).Value
        Dim sglBarcodeWidthScale As Single: sglBarcodeWidthScale = shtData.Range(gstrWidthScaleCellName).Value
        Dim sglMaxColumn As Single: sglMaxColumn = shtData.Range(gstrColumnsOfBarcodes).Value
        Dim intSymbology As Integer: intSymbology = shtData.Range(gstrSymbologyCellName).Value
        Dim sglTextPosition As Single: sglTextPosition = shtData.Range(gstrTextPositionCellName).Value
        Dim sglTextboxHeight As Single: sglTextboxHeight = shtData.Range(gstrTextBoxHeightCellName).Value
        Dim sglBarcodeLeft As Single: sglBarcodeLeft = 0
        Dim intBarcodeTop As Single: intBarcodeTop = 0
        Dim intCurrentColumn As Integer: intCurrentColumn = 1 ' will be set to 1 anyway
        Dim intCurrentRow As Integer:  intCurrentRow = 1 ' will be set to 1 anyway
        Dim strEncoded  As String
        For Each cl In rngData.Cells
            If (intCurrentColumn Mod (sglMaxColumn + 1)) = 0 Then
                If intCurrentRow = 1 And intCurrentColumn = 1 Then ' First row
                    intBarcodeTop = sglTopPadding
                Else 'Next row
                    intCurrentRow = intCurrentRow + 1
                    intBarcodeTop = intBarcodeTop + sglTopPadding + sglBarcodeHeight + sglBottomPadding
                End If
                intCurrentColumn = 1
                sglBarcodeLeft = sglLeftPadding
            Else 'Same row, next column
                If intCurrentColumn = 1 Then
                    sglBarcodeLeft = sglLeftPadding
                    If intCurrentRow = 1 Then
                        intBarcodeTop = sglTopPadding
                    End If
                Else ' (not first column)
                    sglBarcodeLeft = sglBarcodeLeft + sglRightPadding + sglLeftPadding + shp.Width ' this is the previous shape's width
                End If
            End If
            Select Case intSymbology
                Case 1
                    strEncoded = Code128B(cl.Value)
                Case 2
                    strEncoded = CodeEAN128_1(cl.Value)
                Case Else ' Encode to the default
                    strEncoded = Code128B(cl.Value)
            End Select
            Set shp = DrawBarcode( _
                strEncoded, _
                sglBarcodeLeft, _
                intBarcodeTop, _
                1, _
                sglBarcodeHeight, _
                30, _
                shtOutput)
            shp.ScaleWidth sglBarcodeWidthScale, msoFalse, msoScaleFromTopLeft
            If sglFontSize > 0 Then
                '-------------------------------------------------------------------
                'Place the string below and centered to the group...
                Set shpText = shtOutput.Shapes.AddTextbox( _
                    msoTextOrientationHorizontal, _
                    shp.Left + (shp.Width / 2), _
                    shp.Top + shp.Height + sglTextPosition, _
                    shp.Width, _
                    sglFontSize _
                )
                With shpText
                    .IncrementLeft -1 * (shpText.Width / 2)
                    .Title = cl.Value
                    .AlternativeText = cl.Value
                    .TextEffect.Text = cl.Value
                    .TextEffect.Alignment = msoTextEffectAlignmentCentered
                    .DrawingObject.Border.LineStyle = xlLineStyleNone
                    .TextFrame.AutoSize = True
                    .TextEffect.FontSize = sglFontSize
                    .TextEffect.FontName = shtData.Range(gstrFontNameCellName).Value
                    .TextFrame.MarginBottom = 0
                    .TextFrame.MarginLeft = 0
                    .TextFrame.MarginRight = 0
                    .TextFrame.MarginTop = 0
                    If sglTextboxHeight < sglFontSize Then
                        .DrawingObject.Height = sglFontSize
                    Else
                    .DrawingObject.Height = sglTextboxHeight
                    End If
                End With
                shp.ZOrder msoSendToBack
                'Re-Group of barcode should include the textbox
                shtOutput.Activate
                shtOutput.Shapes.Range(Array(shp.Name, shpText.Name)).Select
                Selection.ShapeRange.Group
            End If
            intCurrentColumn = intCurrentColumn + 1
        Next cl
    End If
    Set rngData = Nothing
    Set shtData = Nothing
    shtOutput.Activate
    SetEcho True
End Sub

Sub mDeleteAllShapes(wsh As Worksheet)
    Dim sp As Shape
    For Each sp In wsh.Shapes
        sp.Delete
    Next
End Sub

Public Function DrawBarcode( _
    EncStr As String, _
    Left As Single, _
    Top As Single, _
    SingleWidth As Single, _
    Height As Single, _
    Optional Color As Long, _
    Optional TargetSheet As Worksheet _
) As Shape

If TargetSheet Is Nothing Then
    Set TargetSheet = ThisWorkbook.ActiveSheet
End If

'
' Parameters:
'
' EncStr - a string of ones and zeros, e.g., "11001011"
' Left - the position (in points) of the upper-left corner of the barcode
'        relative to the upper-left corner of the worksheet.
' Top - the position (in points) of the upper-left corner of the barcode
'       relative to the upper-left corner of the worksheet.
' SingleWidth - the width (in points) of a single-wide bar or space.
' Height - the height of the bars, in points.
' Color - (optional) the color of bars; if omitted, the color vill be black.
'
Dim Bars() As BarParams
Dim NextBar As Boolean
Dim i, j As Long
Dim BarColl() As Variant
'
ReDim Bars(1 To 1)
Bars(1).Width = 0
NextBar = False
j = 1
'
For i = 1 To Len(EncStr) Step 1
    If Mid(EncStr, i, 1) = "1" Then
        If Not NextBar Then Bars(j).Pos = i
        Bars(j).Width = Bars(j).Width + 1
        NextBar = True
    Else
        If NextBar Then
            j = j + 1
            ReDim Preserve Bars(1 To j)
            Bars(j).Width = 0
        End If
        NextBar = False
    End If
Next i
'
ReDim BarColl(1 To j)
'
For i = 1 To j Step 1
    With TargetSheet.Shapes.AddShape(msoShapeRectangle, _
        Left + (Bars(i).Pos - 1) * SingleWidth, Top, _
        Bars(i).Width * SingleWidth, Height)
        .Line.Visible = msoFalse
        .Fill.ForeColor.RGB = Color
        BarColl(i) = .Name
    End With
Next i
' [TODO] remove next two dims
Dim shpGroup As GroupShapes
Dim test As Variant
Set DrawBarcode = TargetSheet.Shapes.Range(BarColl).Group
End Function

Function Code128B(TxtStr As String) As String
'
' Parameters
'
' TxtSrt - an alphanumeric string; Chr(32) to Chr(106) can be used.
'
Const MaxChB = 94
'
Dim i, j As Long
Dim SymChB(0 To MaxChB) As String * 1
Dim SymEnc As Variant
Dim WgtSum As Long
Dim EncStr As String
'
For i = 0 To 94
    SymChB(i) = Chr(i + 32)
Next i
'
SymEnc = Array( _
    "11011001100", "11001101100", "11001100110", "10010011000", "10010001100", _
    "10001001100", "10011001000", "10011000100", "10001100100", "11001001000", _
    "11001000100", "11000100100", "10110011100", "10011011100", "10011001110", _
    "10111001100", "10011101100", "10011100110", "11001110010", "11001011100", _
    "11001001110", "11011100100", "11001110100", "11101101110", "11101001100", _
    "11100101100", "11100100110", "11101100100", "11100110100", "11100110010", _
    "11011011000", "11011000110", "11000110110", "10100011000", "10001011000", _
    "10001000110", "10110001000", "10001101000", "10001100010", "11010001000", _
    "11000101000", "11000100010", "10110111000", "10110001110", "10001101110", _
    "10111011000", "10111000110", "10001110110", "11101110110", "11010001110", _
    "11000101110", "11011101000", "11011100010", "11011101110", "11101011000", _
    "11101000110", "11100010110", "11101101000", "11101100010", "11100011010", _
    "11101111010", "11001000010", "11110001010", "10100110000", "10100001100", _
    "10010110000", "10010000110", "10000101100", "10000100110", "10110010000", _
    "10110000100", "10011010000", "10011000010", "10000110100", "10000110010", _
    "11000010010", "11001010000", "11110111010", "11000010100", "10001111010", _
    "10100111100", "10010111100", "10010011110", "10111100100", "10011110100", _
    "10011110010", "11110100100", "11110010100", "11110010010", "11011011110", _
    "11011110110", "11110110110", "10101111000", "10100011110", "10001011110", _
    "10111101000", "10111100010", "11110101000", "11110100010", "10111011110", _
    "10111101110", "11101011110", "11110101110", "11010000100", "11010010000", _
    "11010011100", "11000111010")
ReDim Preserve SymEnc(0 To 106)
'
WgtSum = 104 ' START-B
EncStr = EncStr + SymEnc(104)
For i = 1 To Len(TxtStr) Step 1
    j = 0
    Do While (Mid(TxtStr, i, 1) <> SymChB(j)) And (j <= MaxChB)
        j = j + 1
    Loop
    If j > MaxChB Then j = 0
    WgtSum = WgtSum + i * j
    EncStr = EncStr + SymEnc(j)
Next i
Code128B = EncStr + SymEnc(WgtSum Mod 103) + SymEnc(106) + "11"
'
End Function


Function CodeEAN128_1(TxtStr As String) As String
'
' Parameters
'
' TxtSrt - an alphanumeric string; Chr(32) to Chr(106) can be used.
'
Const MaxChB = 94
'
Dim i, j As Long
Dim SymChB(0 To MaxChB) As String * 1
Dim SymEnc As Variant
Dim WgtSum As Long
Dim EncStr As String
'
For i = 0 To 94
    SymChB(i) = Chr(i + 32)
Next i
'
SymEnc = Array( _
    "11011001100", "11001101100", "11001100110", "10010011000", "10010001100", _
    "10001001100", "10011001000", "10011000100", "10001100100", "11001001000", _
    "11001000100", "11000100100", "10110011100", "10011011100", "10011001110", _
    "10111001100", "10011101100", "10011100110", "11001110010", "11001011100", _
    "11001001110", "11011100100", "11001110100", "11101101110", "11101001100", _
    "11100101100", "11100100110", "11101100100", "11100110100", "11100110010", _
    "11011011000", "11011000110", "11000110110", "10100011000", "10001011000", _
    "10001000110", "10110001000", "10001101000", "10001100010", "11010001000", _
    "11000101000", "11000100010", "10110111000", "10110001110", "10001101110", _
    "10111011000", "10111000110", "10001110110", "11101110110", "11010001110", _
    "11000101110", "11011101000", "11011100010", "11011101110", "11101011000", _
    "11101000110", "11100010110", "11101101000", "11101100010", "11100011010", _
    "11101111010", "11001000010", "11110001010", "10100110000", "10100001100", _
    "10010110000", "10010000110", "10000101100", "10000100110", "10110010000", _
    "10110000100", "10011010000", "10011000010", "10000110100", "10000110010", _
    "11000010010", "11001010000", "11110111010", "11000010100", "10001111010", _
    "10100111100", "10010111100", "10010011110", "10111100100", "10011110100", _
    "10011110010", "11110100100", "11110010100", "11110010010", "11011011110", _
    "11011110110", "11110110110", "10101111000", "10100011110", "10001011110", _
    "10111101000", "10111100010", "11110101000", "11110100010", "10111011110", _
    "10111101110", "11101011110", "11110101110", "11010000100", "11010010000", _
    "11010011100", "11000111010")
ReDim Preserve SymEnc(0 To 106)
'
WgtSum = 104 + 102 ' START-B, FNC1
EncStr = EncStr + SymEnc(104) + SymEnc(102)
For i = 1 To Len(TxtStr) Step 1
    j = 0
    Do While (Mid(TxtStr, i, 1) <> SymChB(j)) And (j <= MaxChB)
        j = j + 1
    Loop
    If j > MaxChB Then j = 0
    WgtSum = WgtSum + (i + 1) * j
    EncStr = EncStr + SymEnc(j)
Next i
CodeEAN128_1 = EncStr + SymEnc(WgtSum Mod 103) + SymEnc(106) + "11"
'
End Function

'=============================================================================================
'From https://www.mrexcel.com/forum/excel-questions/784030-code128-barcode-generator-vba.html
'=============================================================================================
Sub Code128Generate_v2(ByVal X As Single, ByVal Y As Single, ByVal Height As Single, ByVal LineWeight As Single, _
                  ByRef TargetSheet As Worksheet, ByVal Content As String, Optional MaxWidth As Single = 0)
' Supports B and C charsets only; values 00-94, 99,101, 103-105 for B, 00-101, 103-105 for C
' X in mm (0.351)
' Y in mm (0.351) 1mm = 2.8 pt
' Height in mm
' LineWeight in pt

Dim WeightSum As Single
Const XmmTopt As Single = 0.351
Const YmmTopt As Single = 0.351
Const XCompRatio As Single = 0.9

Const Tbar_Symbol As String * 2 = "11"
Dim CurBar As Integer
Dim i, j, k, CharIndex, SymbolIndex As Integer
Dim tstr2 As String * 2
Dim tstr1 As String * 1
Dim ContentString As String ' bars sequence
Const Asw As String * 1 = "A" ' alpha switch
Const Dsw As String * 1 = "D" 'digital switch
Const Arrdim As Byte = 30

Dim Sw, PrevSw As String * 1  ' switch
Dim BlockIndex, BlockCount, DBlockMod2, DBlockLen As Byte

Dim BlockLen(Arrdim) As Byte
Dim BlockSw(Arrdim) As String * 1

Dim SymbolValue(0 To 106) As Integer ' values
Dim SymbolString(0 To 106) As String * 11 'bits sequence
Dim SymbolCharB(0 To 106) As String * 1  'Chars in B set
Dim SymbolCharC(0 To 106) As String * 2  'Chars in B set

For i = 0 To 106 ' values
    SymbolValue(i) = i
Next i

' Symbols in charset B
For i = 0 To 94
    SymbolCharB(i) = Chr(i + 32)
Next i

' Symbols in charset C
SymbolCharC(0) = "00"
SymbolCharC(1) = "01"
SymbolCharC(2) = "02"
SymbolCharC(3) = "03"
SymbolCharC(4) = "04"
SymbolCharC(5) = "05"
SymbolCharC(6) = "06"
SymbolCharC(7) = "07"
SymbolCharC(8) = "08"
SymbolCharC(9) = "09"
For i = 10 To 99
    SymbolCharC(i) = CStr(i)
Next i

' bit sequences
SymbolString(0) = "11011001100"
SymbolString(1) = "11001101100"
SymbolString(2) = "11001100110"
SymbolString(3) = "10010011000"
SymbolString(4) = "10010001100"
SymbolString(5) = "10001001100"
SymbolString(6) = "10011001000"
SymbolString(7) = "10011000100"
SymbolString(8) = "10001100100"
SymbolString(9) = "11001001000"
SymbolString(10) = "11001000100"
SymbolString(11) = "11000100100"
SymbolString(12) = "10110011100"
SymbolString(13) = "10011011100"
SymbolString(14) = "10011001110"
SymbolString(15) = "10111001100"
SymbolString(16) = "10011101100"
SymbolString(17) = "10011100110"
SymbolString(18) = "11001110010"
SymbolString(19) = "11001011100"
SymbolString(20) = "11001001110"
SymbolString(21) = "11011100100"
SymbolString(22) = "11001110100"
SymbolString(23) = "11101101110"
SymbolString(24) = "11101001100"
SymbolString(25) = "11100101100"
SymbolString(26) = "11100100110"
SymbolString(27) = "11101100100"
SymbolString(28) = "11100110100"
SymbolString(29) = "11100110010"
SymbolString(30) = "11011011000"
SymbolString(31) = "11011000110"
SymbolString(32) = "11000110110"
SymbolString(33) = "10100011000"
SymbolString(34) = "10001011000"
SymbolString(35) = "10001000110"
SymbolString(36) = "10110001000"
SymbolString(37) = "10001101000"
SymbolString(38) = "10001100010"
SymbolString(39) = "11010001000"
SymbolString(40) = "11000101000"
SymbolString(41) = "11000100010"
SymbolString(42) = "10110111000"
SymbolString(43) = "10110001110"
SymbolString(44) = "10001101110"
SymbolString(45) = "10111011000"
SymbolString(46) = "10111000110"
SymbolString(47) = "10001110110"
SymbolString(48) = "11101110110"
SymbolString(49) = "11010001110"
SymbolString(50) = "11000101110"
SymbolString(51) = "11011101000"
SymbolString(52) = "11011100010"
SymbolString(53) = "11011101110"
SymbolString(54) = "11101011000"
SymbolString(55) = "11101000110"
SymbolString(56) = "11100010110"
SymbolString(57) = "11101101000"
SymbolString(58) = "11101100010"
SymbolString(59) = "11100011010"
SymbolString(60) = "11101111010"
SymbolString(61) = "11001000010"
SymbolString(62) = "11110001010"
SymbolString(63) = "10100110000"
SymbolString(64) = "10100001100"
SymbolString(65) = "10010110000"
SymbolString(66) = "10010000110"
SymbolString(67) = "10000101100"
SymbolString(68) = "10000100110"
SymbolString(69) = "10110010000"
SymbolString(70) = "10110000100"
SymbolString(71) = "10011010000"
SymbolString(72) = "10011000010"
SymbolString(73) = "10000110100"
SymbolString(74) = "10000110010"
SymbolString(75) = "11000010010"
SymbolString(76) = "11001010000"
SymbolString(77) = "11110111010"
SymbolString(78) = "11000010100"
SymbolString(79) = "10001111010"
SymbolString(80) = "10100111100"
SymbolString(81) = "10010111100"
SymbolString(82) = "10010011110"
SymbolString(83) = "10111100100"
SymbolString(84) = "10011110100"
SymbolString(85) = "10011110010"
SymbolString(86) = "11110100100"
SymbolString(87) = "11110010100"
SymbolString(88) = "11110010010"
SymbolString(89) = "11011011110"
SymbolString(90) = "11011110110"
SymbolString(91) = "11110110110"
SymbolString(92) = "10101111000"
SymbolString(93) = "10100011110"
SymbolString(94) = "10001011110"
SymbolString(95) = "10111101000"
SymbolString(96) = "10111100010"
SymbolString(97) = "11110101000"
SymbolString(98) = "11110100010"
SymbolString(99) = "10111011110"
SymbolString(100) = "10111101110"
SymbolString(101) = "11101011110"
SymbolString(102) = "11110101110"
SymbolString(103) = "11010000100"
SymbolString(104) = "11010010000"
SymbolString(105) = "11010011100"
SymbolString(106) = "11000111010"


X = X / XmmTopt 'mm to pt
Y = Y / YmmTopt 'mm to pt
Height = Height / YmmTopt 'mm to pt


If IsNumeric(Content) = True And Len(Content) Mod 2 = 0 Then 'numeric, mode C
   WeightSum = SymbolValue(105) ' start-c
   ContentString = ContentString + SymbolString(105)
   i = 0 ' symbol count
   For j = 1 To Len(Content) Step 2
      tstr2 = Mid(Content, j, 2)
      i = i + 1
      k = 0
      Do While tstr2 <> SymbolCharC(k)
         k = k + 1
      Loop
      WeightSum = WeightSum + i * SymbolValue(k)
      ContentString = ContentString + SymbolString(k)
   Next j
   ContentString = ContentString + SymbolString(SymbolValue(WeightSum Mod 103))
   ContentString = ContentString + SymbolString(106)
   ContentString = ContentString + Tbar_Symbol
   
Else ' alpha-numeric
   
   ' first digit
   Select Case IsNumeric(Mid(Content, 1, 1))
   Case Is = True 'digit
      Sw = Dsw
   Case Is = False 'alpha
      Sw = Asw
   End Select
   BlockCount = 1
   BlockSw(BlockCount) = Sw
   BlockIndex = 1
   BlockLen(BlockCount) = 1 'block length

   i = 2 ' symbol index
   
   Do While i <= Len(Content)
      Select Case IsNumeric(Mid(Content, i, 1))
      Case Is = True 'digit
         Sw = Dsw
      Case Is = False 'alpha
         Sw = Asw
      End Select
      
      If Sw = BlockSw(BlockCount) Then
         BlockLen(BlockCount) = BlockLen(BlockCount) + 1
      Else
         BlockCount = BlockCount + 1
         BlockSw(BlockCount) = Sw
         BlockLen(BlockCount) = 1
         BlockIndex = BlockIndex + 1

      End If
      
      i = i + 1
   Loop

   'encoding
   CharIndex = 1 'index of Content character
   SymbolIndex = 0
   
   For BlockIndex = 1 To BlockCount ' encoding by blocks

      If BlockSw(BlockIndex) = Dsw And BlockLen(BlockIndex) >= 4 Then ' switch to C
         Select Case BlockIndex
         Case Is = 1
            WeightSum = SymbolValue(105) ' Start-C
            ContentString = ContentString + SymbolString(105)
         Case Else
            SymbolIndex = SymbolIndex + 1
            WeightSum = WeightSum + SymbolIndex * SymbolValue(99) 'switch c
            ContentString = ContentString + SymbolString(99)
         End Select
         PrevSw = Dsw
         
         ' encoding even amount of chars in a D block
         DBlockMod2 = BlockLen(BlockIndex) Mod 2
         If DBlockMod2 <> 0 Then 'even chars always to encode
            DBlockLen = BlockLen(BlockIndex) - DBlockMod2
         Else
            DBlockLen = BlockLen(BlockIndex)
         End If
         
         For j = 1 To DBlockLen / 2 Step 1
            tstr2 = Mid(Content, CharIndex, 2)
            CharIndex = CharIndex + 2
            SymbolIndex = SymbolIndex + 1
            k = 0
            Do While tstr2 <> SymbolCharC(k)
               k = k + 1
            Loop
            WeightSum = WeightSum + SymbolIndex * SymbolValue(k)
            ContentString = ContentString + SymbolString(k)
         Next j
         
         If DBlockMod2 <> 0 Then ' switch to B, encode 1 char
            PrevSw = Asw
            SymbolIndex = SymbolIndex + 1
            WeightSum = WeightSum + SymbolIndex * SymbolValue(100) 'switch b
            ContentString = ContentString + SymbolString(100)
            
            'CharIndex = CharIndex + 1
            SymbolIndex = SymbolIndex + 1
            tstr1 = Mid(Content, CharIndex, 1)
            k = 0
            Do While tstr1 <> SymbolCharB(k)
               k = k + 1
            Loop
            WeightSum = WeightSum + SymbolIndex * SymbolValue(k)
            ContentString = ContentString + SymbolString(k)
         End If
         
      Else 'alpha in B mode
         Select Case BlockIndex
         Case Is = 1
         '   PrevSw = Asw
            WeightSum = SymbolValue(104) ' start-b
            ContentString = ContentString + SymbolString(104)
         Case Else
            If PrevSw <> Asw Then
               SymbolIndex = SymbolIndex + 1
               WeightSum = WeightSum + SymbolIndex * SymbolValue(100) 'switch b
               ContentString = ContentString + SymbolString(100)
               
            End If
         End Select
         PrevSw = Asw
         
         For j = CharIndex To CharIndex + BlockLen(BlockIndex) - 1 Step 1
            tstr1 = Mid(Content, j, 1)
            SymbolIndex = SymbolIndex + 1
            k = 0
            Do While tstr1 <> SymbolCharB(k)
               k = k + 1
            Loop
            WeightSum = WeightSum + SymbolIndex * SymbolValue(k)
            ContentString = ContentString + SymbolString(k)
         Next j
         CharIndex = j


      End If
   Next BlockIndex
   ContentString = ContentString + SymbolString(SymbolValue(WeightSum Mod 103))
   ContentString = ContentString + SymbolString(106)
   ContentString = ContentString + Tbar_Symbol
   
End If

   If MaxWidth > 0 And Len(ContentString) * LineWeight * XmmTopt > MaxWidth Then
      LineWeight = MaxWidth / (Len(ContentString) * XmmTopt)
      LineWeight = LineWeight / XCompRatio
   End If
   
'Barcode drawing
CurBar = 0

For i = 1 To Len(ContentString)
    Select Case Mid(ContentString, i, 1)
    Case 0
        CurBar = CurBar + 1
    Case 1
        CurBar = CurBar + 1
        With TargetSheet.Shapes.AddLine(X + (CurBar * LineWeight) * XCompRatio, Y, X + (CurBar * LineWeight) * XCompRatio, (Y + Height)).Line
        .Weight = LineWeight
        .ForeColor.RGB = vbBlack
        End With
    End Select
Next i

End Sub

'From https://www.mrexcel.com/forum/excel-questions/784030-code128-barcode-generator-vba.html
Sub Code128GenerateLegacy(ByVal X As Single, ByVal Y As Single, ByVal Height As Single, ByVal LineWeight As Single, _
                  ByRef TargetSheet As Worksheet, ByVal Content As String)
' Supports B and C charsets only; values 00-94, 99,101, 103-105 for B, 00-101, 103-105 for C
' X, Y - top-left corner coordinates
' X in mm (0.376042)
' Y in mm (0.341)
' Height in mm
' LineWeight in pt

Const Tbar_Symbol As String * 2 = "11" ' termination bar
Dim WeightSum As Single
Dim CurBar As Integer
Dim i, j, k, FirstSymbol As Integer
Dim tstr2 As String * 2
Dim tstr1 As String * 1
Dim ContentString As String ' bars sequence

Dim SymbolValue(0 To 106) As Integer ' values
Dim SymbolString(0 To 106) As String * 11 'bits sequence
Dim SymbolCharB(0 To 106) As String * 1  'Chars in B set
Dim SymbolCharC(0 To 106) As String * 2  'Chars in B set

For i = 0 To 106 ' values
    SymbolValue(i) = i
Next i

' Symbols in charset B
For i = 0 To 94
    SymbolCharB(i) = Chr(i + 32)
Next i

' Symbols in charset C
SymbolCharC(0) = "00"
SymbolCharC(1) = "01"
SymbolCharC(2) = "02"
SymbolCharC(3) = "03"
SymbolCharC(4) = "04"
SymbolCharC(5) = "05"
SymbolCharC(6) = "06"
SymbolCharC(7) = "07"
SymbolCharC(8) = "08"
SymbolCharC(9) = "09"
For i = 10 To 99
    SymbolCharC(i) = CStr(i)
Next i

' bit sequences
SymbolString(0) = "11011001100"
SymbolString(1) = "11001101100"
SymbolString(2) = "11001100110"
SymbolString(3) = "10010011000"
SymbolString(4) = "10010001100"
SymbolString(5) = "10001001100"
SymbolString(6) = "10011001000"
SymbolString(7) = "10011000100"
SymbolString(8) = "10001100100"
SymbolString(9) = "11001001000"
SymbolString(10) = "11001000100"
SymbolString(11) = "11000100100"
SymbolString(12) = "10110011100"
SymbolString(13) = "10011011100"
SymbolString(14) = "10011001110"
SymbolString(15) = "10111001100"
SymbolString(16) = "10011101100"
SymbolString(17) = "10011100110"
SymbolString(18) = "11001110010"
SymbolString(19) = "11001011100"
SymbolString(20) = "11001001110"
SymbolString(21) = "11011100100"
SymbolString(22) = "11001110100"
SymbolString(23) = "11101101110"
SymbolString(24) = "11101001100"
SymbolString(25) = "11100101100"
SymbolString(26) = "11100100110"
SymbolString(27) = "11101100100"
SymbolString(28) = "11100110100"
SymbolString(29) = "11100110010"
SymbolString(30) = "11011011000"
SymbolString(31) = "11011000110"
SymbolString(32) = "11000110110"
SymbolString(33) = "10100011000"
SymbolString(34) = "10001011000"
SymbolString(35) = "10001000110"
SymbolString(36) = "10110001000"
SymbolString(37) = "10001101000"
SymbolString(38) = "10001100010"
SymbolString(39) = "11010001000"
SymbolString(40) = "11000101000"
SymbolString(41) = "11000100010"
SymbolString(42) = "10110111000"
SymbolString(43) = "10110001110"
SymbolString(44) = "10001101110"
SymbolString(45) = "10111011000"
SymbolString(46) = "10111000110"
SymbolString(47) = "10001110110"
SymbolString(48) = "11101110110"
SymbolString(49) = "11010001110"
SymbolString(50) = "11000101110"
SymbolString(51) = "11011101000"
SymbolString(52) = "11011100010"
SymbolString(53) = "11011101110"
SymbolString(54) = "11101011000"
SymbolString(55) = "11101000110"
SymbolString(56) = "11100010110"
SymbolString(57) = "11101101000"
SymbolString(58) = "11101100010"
SymbolString(59) = "11100011010"
SymbolString(60) = "11101111010"
SymbolString(61) = "11001000010"
SymbolString(62) = "11110001010"
SymbolString(63) = "10100110000"
SymbolString(64) = "10100001100"
SymbolString(65) = "10010110000"
SymbolString(66) = "10010000110"
SymbolString(67) = "10000101100"
SymbolString(68) = "10000100110"
SymbolString(69) = "10110010000"
SymbolString(70) = "10110000100"
SymbolString(71) = "10011010000"
SymbolString(72) = "10011000010"
SymbolString(73) = "10000110100"
SymbolString(74) = "10000110010"
SymbolString(75) = "11000010010"
SymbolString(76) = "11001010000"
SymbolString(77) = "11110111010"
SymbolString(78) = "11000010100"
SymbolString(79) = "10001111010"
SymbolString(80) = "10100111100"
SymbolString(81) = "10010111100"
SymbolString(82) = "10010011110"
SymbolString(83) = "10111100100"
SymbolString(84) = "10011110100"
SymbolString(85) = "10011110010"
SymbolString(86) = "11110100100"
SymbolString(87) = "11110010100"
SymbolString(88) = "11110010010"
SymbolString(89) = "11011011110"
SymbolString(90) = "11011110110"
SymbolString(91) = "11110110110"
SymbolString(92) = "10101111000"
SymbolString(93) = "10100011110"
SymbolString(94) = "10001011110"
SymbolString(95) = "10111101000"
SymbolString(96) = "10111100010"
SymbolString(97) = "11110101000"
SymbolString(98) = "11110100010"
SymbolString(99) = "10111011110"
SymbolString(100) = "10111101110"
SymbolString(101) = "11101011110"
SymbolString(102) = "11110101110"
SymbolString(103) = "11010000100"
SymbolString(104) = "11010010000"
SymbolString(105) = "11010011100"
SymbolString(106) = "11000111010"

X = X / 0.376042 'mm to pt
Y = Y / 0.341 'mm to pt
Height = Height / 0.341 'mm to pt

If IsNumeric(Content) = True Then  ' value is numeric
   i = 1 'symbol and weight index
   If Len(Content) Mod 2 = 1 Then 'odd
       WeightSum = SymbolValue(104) ' start-b
       ContentString = ContentString + SymbolString(104)
      tstr1 = Mid(Content, 1, 1)
      k = 0
      Do While tstr1 <> SymbolCharB(k)
         k = k + 1
      Loop
      WeightSum = WeightSum + i * SymbolValue(k)
      ContentString = ContentString + SymbolString(k)
      i = i + 1
      WeightSum = WeightSum + i * SymbolValue(99) 'Code-C
      ContentString = ContentString + SymbolString(99) 'Code-C
      Content = Right(Content, Len(Content) - 1) 'cut 1st symbol
   Else 'even
      WeightSum = SymbolValue(105) ' start-c
      ContentString = ContentString + SymbolString(105)
      i = 0
   End If
   
   For j = 1 To Len(Content) Step 2
      tstr2 = Mid(Content, j, 2)
      i = i + 1
      k = 0
      Do While tstr2 <> SymbolCharC(k)
         k = k + 1
      Loop
      WeightSum = WeightSum + i * SymbolValue(k)
      ContentString = ContentString + SymbolString(k)
   Next j
   ContentString = ContentString + SymbolString(SymbolValue(WeightSum Mod 103))
   ContentString = ContentString + SymbolString(106)

   
   Else ' alpha-numeric
   WeightSum = SymbolValue(104) ' start-b
   ContentString = ContentString + SymbolString(104)
   i = 0 ' symbol count
   For j = 1 To Len(Content) Step 1
      tstr1 = Mid(Content, j, 1)
      i = i + 1
      k = 0
      Do While tstr1 <> SymbolCharB(k)
         k = k + 1
      Loop
      WeightSum = WeightSum + i * SymbolValue(k)
      ContentString = ContentString + SymbolString(k)
   Next j
   ContentString = ContentString + SymbolString(SymbolValue(WeightSum Mod 103))
   ContentString = ContentString + SymbolString(106)

End If

ContentString = ContentString + Tbar_Symbol

'Barcode drawing
CurBar = 0

For i = 1 To Len(ContentString)
    Select Case Mid(ContentString, i, 1)
    Case 0
        CurBar = CurBar + 1
    Case 1
        CurBar = CurBar + 1
' (CurBar * LineWeight) * 0.9 -  here is 10% overlapping :-)
        With TargetSheet.Shapes.AddLine(X + (CurBar * LineWeight) * 0.9, Y, X + (CurBar * LineWeight) * 0.9, (Y + Height)).Line
        .Weight = LineWeight
        .ForeColor.RGB = vbBlack ' my Excel writes light-blue lines by default, so the color is forcibly switched
        End With
    End Select
Next i

End Sub
