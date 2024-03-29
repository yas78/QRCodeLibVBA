VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Symbol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const DEFAULT_MODULE_SIZE As Long = 5
Private Const MIN_MODULE_SIZE     As Long = 2

Private m_parent As Symbols

Private m_position As Long

Private m_currEncoder      As IQRCodeEncoder
Private m_currEncodingMode As EncodingMode
Private m_currVersion      As Long

Private m_dataBitCapacity As Long
Private m_dataBitCounter  As Long

Private m_segments As Collection
Private m_segmentCounter(EncodingMode.NUMERIC To EncodingMode.KANJI) As Long

Private m_index As Long

Friend Sub Init(ByVal parentObj As Symbols)
    Set m_parent = parentObj
    m_index = parentObj.Count

    m_position = parentObj.Count

    Set m_currEncoder = Nothing
    m_currEncodingMode = EncodingMode.UNKNOWN
    m_currVersion = parentObj.MinVersion

    m_dataBitCapacity = 8 * DataCodeword.GetTotalNumber( _
        parentObj.ErrorCorrectionLevel, parentObj.MinVersion)

    m_dataBitCounter = 0

    Set m_segments = New Collection

    If parentObj.StructuredAppend Then
        m_dataBitCapacity = m_dataBitCapacity - StructuredAppend.HEADER_LENGTH
    End If
End Sub

Friend Property Get Parent() As Symbols
    Set Parent = m_parent
End Property

Friend Property Get Index() As Long
    Index = m_index
End Property

Friend Property Get Version() As Long
    Version = m_currVersion
End Property

Friend Property Get CurrentEncodingMode() As EncodingMode
    CurrentEncodingMode = m_currEncodingMode
End Property

Friend Function TryAppend(ByVal c As String) As Boolean
    Dim bitLength As Long
    bitLength = m_currEncoder.GetCodewordBitLength(c)

    Do While (m_dataBitCapacity < m_dataBitCounter + bitLength)
        If m_currVersion >= m_parent.MaxVersion Then
            TryAppend = False
            Exit Function
        End If

        Call SelectVersion
    Loop

    Call m_currEncoder.Append(c)
    m_dataBitCounter = m_dataBitCounter + bitLength
    Call m_parent.UpdateParity(c)

    TryAppend = True
End Function

Friend Function TrySetEncodingMode(ByVal encMode As EncodingMode, _
                                   ByVal c As String) As Boolean
    Dim enc As IQRCodeEncoder

    Select Case encMode
        Case EncodingMode.NUMERIC
            Set enc = New NumericEncoder
        Case EncodingMode.ALPHA_NUMERIC
            Set enc = New AlphanumericEncoder
        Case EncodingMode.EIGHT_BIT_BYTE
            Set enc = New ByteEncoder
            Call enc.Init(m_parent.Encoding)
        Case EncodingMode.KANJI
            If Charset.IsJP(m_parent.Encoding.Charset) Then
                Set enc = New KanjiEncoder
                Call enc.Init(m_parent.Encoding)
            Else
                Call Err.Raise(51)
            End If
        Case Else
            Call Err.Raise(5)
    End Select

    Dim bitLength As Long
    bitLength = enc.GetCodewordBitLength(c)

    Do While (m_dataBitCapacity < _
                m_dataBitCounter + _
                ModeIndicator.Length + _
                CharCountIndicator.GetLength(m_currVersion, encMode) + _
                bitLength)
        If m_currVersion >= m_parent.MaxVersion Then
            TrySetEncodingMode = False
            Exit Function
        End If

        Call SelectVersion
    Loop

    m_dataBitCounter = m_dataBitCounter + _
                       ModeIndicator.Length + _
                       CharCountIndicator.GetLength(m_currVersion, encMode)

    Set m_currEncoder = enc
    Call m_segments.Add(enc)
    m_segmentCounter(encMode) = m_segmentCounter(encMode) + 1
    m_currEncodingMode = encMode

    TrySetEncodingMode = True
End Function

Private Sub SelectVersion()
    Dim num As Long
    Dim encMode As EncodingMode
    For encMode = EncodingMode.NUMERIC To EncodingMode.KANJI
        num = m_segmentCounter(encMode)
        m_dataBitCounter = m_dataBitCounter + _
                           num * CharCountIndicator.GetLength( _
                                m_currVersion + 1, encMode) - _
                           num * CharCountIndicator.GetLength( _
                                m_currVersion + 0, encMode)
    Next

    m_currVersion = m_currVersion + 1
    m_dataBitCapacity = 8 * DataCodeword.GetTotalNumber( _
        m_parent.ErrorCorrectionLevel, m_currVersion)
    m_parent.MinVersion = m_currVersion

    If m_parent.StructuredAppend Then
        m_dataBitCapacity = m_dataBitCapacity - StructuredAppend.HEADER_LENGTH
    End If
End Sub

Private Function BuildDataBlock() As Variant()
    Dim dataBytes() As Byte
    dataBytes = GetMessageBytes()

    Dim numPreBlocks As Long
    numPreBlocks = RSBlock.GetTotalNumber( _
        m_parent.ErrorCorrectionLevel, m_currVersion, True)

    Dim numFolBlocks As Long
    numFolBlocks = RSBlock.GetTotalNumber( _
        m_parent.ErrorCorrectionLevel, m_currVersion, False)

    Dim ret() As Variant
    ReDim ret(numPreBlocks + numFolBlocks - 1)

    Dim sz As Long
    sz = RSBlock.GetNumberDataCodewords( _
        m_parent.ErrorCorrectionLevel, m_currVersion, True)

    Dim srcIdx As Long
    srcIdx = 0

    Dim data() As Byte

    Dim i As Long
    For i = 0 To numPreBlocks - 1
        ReDim data(sz - 1)
        Call ArrayUtil.Copy(data, 0, dataBytes, srcIdx, sz)
        srcIdx = srcIdx + sz
        ret(i) = data
    Next

    sz = RSBlock.GetNumberDataCodewords( _
        m_parent.ErrorCorrectionLevel, m_currVersion, False)

    For i = numPreBlocks To numPreBlocks + numFolBlocks - 1
        ReDim data(sz - 1)
        Call ArrayUtil.Copy(data, 0, dataBytes, srcIdx, sz)
        srcIdx = srcIdx + sz
        ret(i) = data
    Next

    BuildDataBlock = ret
End Function

Private Function BuildErrorCorrectionBlock(ByRef dataBlock() As Variant) As Variant()
    Dim numECCodewords As Long
    numECCodewords = RSBlock.GetNumberECCodewords( _
        m_parent.ErrorCorrectionLevel, m_currVersion)

    Dim numPreBlocks As Long
    numPreBlocks = RSBlock.GetTotalNumber( _
        m_parent.ErrorCorrectionLevel, m_currVersion, True)

    Dim numFolBlocks As Long
    numFolBlocks = RSBlock.GetTotalNumber( _
        m_parent.ErrorCorrectionLevel, m_currVersion, False)

    Dim ret() As Variant
    ReDim ret(numPreBlocks + numFolBlocks - 1)

    Dim eccDataTmp() As Byte
    ReDim eccDataTmp(numECCodewords - 1)

    Dim idx As Long
    For idx = 0 To UBound(ret)
        ret(idx) = eccDataTmp
    Next

    Dim gp() As Long
    gp = GeneratorPolynomials.Item(numECCodewords)

    Dim eccIdx As Long
    Dim data() As Long
    Dim exp    As Long

    Dim blockIdx As Long
    Dim i As Long
    Dim j As Long
    For blockIdx = 0 To UBound(dataBlock)
        ReDim data(UBound(dataBlock(blockIdx)) + UBound(ret(blockIdx)) + 1)
        eccIdx = UBound(data)

        For i = 0 To UBound(dataBlock(blockIdx))
            data(eccIdx) = dataBlock(blockIdx)(i)
            eccIdx = eccIdx - 1
        Next

        For i = UBound(data) To numECCodewords Step -1
            If data(i) > 0 Then
                exp = GaloisField256.ToExp(data(i))
                eccIdx = i

                For j = UBound(gp) To 0 Step -1
                    data(eccIdx) = data(eccIdx) Xor _
                                   GaloisField256.ToInt((gp(j) + exp) Mod 255)
                    eccIdx = eccIdx - 1
                Next
            End If
        Next

        eccIdx = numECCodewords - 1

        For i = 0 To UBound(ret(blockIdx))
            ret(blockIdx)(i) = data(eccIdx)
            eccIdx = eccIdx - 1
        Next
    Next

    BuildErrorCorrectionBlock = ret
End Function

Private Function GetEncodingRegionBytes() As Byte()
    Dim dataBlock() As Variant
    dataBlock = BuildDataBlock()

    Dim ecBlock() As Variant
    ecBlock = BuildErrorCorrectionBlock(dataBlock)

    Dim numCodewords As Long
    numCodewords = Codeword.GetTotalNumber(m_currVersion)

    Dim numDataCodewords As Long
    numDataCodewords = DataCodeword.GetTotalNumber( _
        m_parent.ErrorCorrectionLevel, m_currVersion)

    Dim ret() As Byte
    ReDim ret(numCodewords - 1)

    Dim r As Long
    Dim c As Long

    Dim n As Long
    n = 0

    Dim idx As Long
    idx = 0
    Do While idx < numDataCodewords
        r = n Mod (UBound(dataBlock) + 1)
        c = n \ (UBound(dataBlock) + 1)

        If c <= UBound(dataBlock(r)) Then
            ret(idx) = dataBlock(r)(c)
            idx = idx + 1
        End If

        n = n + 1
    Loop

    n = 0

    Do While idx < numCodewords
        r = n Mod (UBound(ecBlock) + 1)
        c = n \ (UBound(ecBlock) + 1)

        If c <= UBound(ecBlock(r)) Then
            ret(idx) = ecBlock(r)(c)
            idx = idx + 1
        End If

        n = n + 1
    Loop

    GetEncodingRegionBytes = ret
End Function

Private Function GetMessageBytes() As Byte()
    Dim bs As New BitSequence

    If m_parent.Count > 1 Then
        Call WriteStructuredAppendHeader(bs)
    End If

    Call WriteSegments(bs)
    Call WriteTerminator(bs)
    Call WritePaddingBits(bs)
    Call WritePadCodewords(bs)

    GetMessageBytes = bs.GetBytes()
End Function

Private Sub WriteStructuredAppendHeader(ByVal bs As BitSequence)
    Call bs.Append(ModeIndicator.STRUCTURED_APPEND_VALUE, _
                   ModeIndicator.Length)
    Call bs.Append(m_position, _
                   SymbolSequenceIndicator.POSITION_LENGTH)
    Call bs.Append(m_parent.Count - 1, _
                   SymbolSequenceIndicator.TOTAL_NUMBER_LENGTH)
    Call bs.Append(m_parent.Parity, _
                   StructuredAppend.PARITY_DATA_LENGTH)
End Sub

Private Sub WriteSegments(ByVal bs As BitSequence)
    Dim data() As Byte
    Dim codewordBitLength As Long

    Dim segment As IQRCodeEncoder
    Dim i As Long
    For Each segment In m_segments
        Call bs.Append(segment.ModeIndicator, ModeIndicator.Length)
        Call bs.Append(segment.CharCount, _
                       CharCountIndicator.GetLength( _
                            m_currVersion, segment.EncodingMode))

        data = segment.GetBytes()

        For i = 0 To UBound(data) - 1
            Call bs.Append(data(i), 8)
        Next

        codewordBitLength = segment.BitCount Mod 8

        If codewordBitLength = 0 Then
            codewordBitLength = 8
        End If

        Call bs.Append(data(UBound(data)) \ _
                       2 ^ (8 - codewordBitLength), codewordBitLength)
    Next
End Sub

Private Sub WriteTerminator(ByVal bs As BitSequence)
    Dim terminatorLength As Long
    terminatorLength = m_dataBitCapacity - m_dataBitCounter

    If terminatorLength > ModeIndicator.Length Then
        terminatorLength = ModeIndicator.Length
    End If

    Call bs.Append(ModeIndicator.TERMINATOR_VALUE, terminatorLength)
End Sub

Private Sub WritePaddingBits(ByVal bs As BitSequence)
    If bs.Length Mod 8 > 0 Then
        Call bs.Append(&H0, 8 - (bs.Length Mod 8))
    End If
End Sub

Private Sub WritePadCodewords(ByVal bs As BitSequence)
    Dim numDataCodewords As Long
    numDataCodewords = DataCodeword.GetTotalNumber( _
        m_parent.ErrorCorrectionLevel, m_currVersion)

    Dim flg As Boolean
    flg = True

    Dim cnt As Long
    cnt = ((8 * numDataCodewords) - bs.Length) \ 8

    Dim i As Long
    For i = 1 To cnt
        Call bs.Append(IIf(flg, 236, 17), 8)
        flg = Not flg
    Next
End Sub

Private Function GetModuleMatrix() As Variant()
    Dim numModulesPerSide As Long
    numModulesPerSide = Module.GetNumModulesPerSide(m_currVersion)

    Dim moduleMatrix() As Variant
    ReDim moduleMatrix(numModulesPerSide - 1)

    Dim rowArray() As Long

    Dim i As Long
    For i = 0 To UBound(moduleMatrix)
        ReDim rowArray(numModulesPerSide - 1)
        moduleMatrix(i) = rowArray
    Next

    Call FinderPattern.Place(moduleMatrix)
    Call Separator.Place(moduleMatrix)
    Call TimingPattern.Place(moduleMatrix)

    If m_currVersion >= 2 Then
        Call AlignmentPattern.Place(m_currVersion, moduleMatrix)
    End If

    Call FormatInfo.PlaceTempBlank(moduleMatrix)

    If m_currVersion >= 7 Then
        Call VersionInfo.PlaceTempBlank(moduleMatrix)
    End If

    Call PlaceSymbolChar(moduleMatrix)
    Call RemainderBit.Place(moduleMatrix)

    Call Masking.Apply(m_currVersion, m_parent.ErrorCorrectionLevel, moduleMatrix)

    GetModuleMatrix = QuietZone.Place(moduleMatrix)
End Function

Private Sub PlaceSymbolChar(ByRef moduleMatrix() As Variant)
    Const VAL As Long = Values.WORD

    Dim data() As Byte
    data = GetEncodingRegionBytes()

    Dim r As Long
    r = UBound(moduleMatrix)

    Dim c As Long
    c = UBound(moduleMatrix(0))

    Dim toLeft As Boolean
    toLeft = True

    Dim rowDirection As Long
    rowDirection = -1

    Dim v As Variant
    Dim bitPos As Long
    For Each v In data
        bitPos = 7

        Do While bitPos >= 0
            If moduleMatrix(r)(c) = Values.BLANK Then
                moduleMatrix(r)(c) = IIf((v And 2 ^ bitPos) > 0, VAL, -VAL)
                bitPos = bitPos - 1
            End If

            If toLeft Then
                c = c - 1
            Else
                If (r + rowDirection) < 0 Then
                    r = 0
                    rowDirection = 1
                    c = c - 1

                    If c = 6 Then
                        c = 5
                    End If

                ElseIf ((r + rowDirection) > UBound(moduleMatrix)) Then
                    r = UBound(moduleMatrix)
                    rowDirection = -1
                    c = c - 1

                    If c = 6 Then
                        c = 5
                    End If

                Else
                    r = r + rowDirection
                    c = c + 1
                End If
            End If

            toLeft = Not toLeft
        Loop
    Next
End Sub

Private Function GetMonochromeBmp(ByVal moduleSize As Long, _
                                  ByVal foreColorRgb As Long, _
                                  ByVal backColorRgb As Long) As Byte()
    Dim moduleMatrix() As Variant
    moduleMatrix = GetModuleMatrix()

    Dim moduleCount As Long
    moduleCount = UBound(moduleMatrix) + 1

    Dim pictWidth As Long
    pictWidth = moduleCount * moduleSize

    Dim pictHeight As Long
    pictHeight = pictWidth

    Dim rowBytesLen As Long
    rowBytesLen = (pictWidth + 7) \ 8

    Dim pack8bit As Long
    If pictWidth Mod 8 > 0 Then
        pack8bit = 8 - (pictWidth Mod 8)
    End If

    Dim pack32bit As Long
    If rowBytesLen Mod 4 > 0 Then
        pack32bit = 8 * (4 - (rowBytesLen Mod 4))
    End If

    Dim rowSize As Long
    rowSize = (pictWidth + pack8bit + pack32bit) \ 8

    Dim bitmapData() As Byte
    ReDim bitmapData(rowSize * pictHeight - 1)

    Dim offset As Long
    offset = 0

    Dim bs As New BitSequence

    Dim pixelColor  As Long
    Dim bitmapRow() As Byte

    Dim r As Long
    Dim v As Variant
    Dim i As Long
    For r = UBound(moduleMatrix) To 0 Step -1
        Call bs.Clear

        For Each v In moduleMatrix(r)
            pixelColor = IIf(Values.IsDark(v), 0, 1)

            For i = 1 To moduleSize
                Call bs.Append(pixelColor, 1)
            Next
        Next
        Call bs.Append(0, pack8bit)
        Call bs.Append(0, pack32bit)
        bitmapRow = bs.GetBytes()

        For i = 1 To moduleSize
            offset = ArrayUtil.CopyAll(bitmapData, offset, bitmapRow)
        Next
    Next

    GetMonochromeBmp = Dib.GetDib(bitmapData, pictWidth, pictHeight, foreColorRgb, backColorRgb, True)
End Function

Private Function GetTrueColorBmp(ByVal moduleSize As Long, _
                                 ByVal foreColorRgb As Long, _
                                 ByVal backColorRgb As Long) As Byte()
    Dim moduleMatrix() As Variant
    moduleMatrix = GetModuleMatrix()

    Dim pictWidth As Long
    pictWidth = (UBound(moduleMatrix(0)) + 1) * moduleSize

    Dim pictHeight As Long
    pictHeight = pictWidth

    Dim rowSize As Long
    rowSize = ((3 * pictWidth + 3) \ 4) * 4

    Dim foreRgb() As Byte
    foreRgb = BitConverter.GetBytes(foreColorRgb, True)

    Dim backRgb() As Byte
    backRgb = BitConverter.GetBytes(backColorRgb, True)

    Dim colorRgb()  As Byte
    Dim bitmapRow() As Byte
    Dim idx         As Long

    Dim bs As New ByteSequence

    Dim r As Long
    Dim v As Variant
    Dim i As Long
    For r = UBound(moduleMatrix) To 0 Step -1
        ReDim bitmapRow(rowSize - 1)
        idx = 0

        For Each v In moduleMatrix(r)
            colorRgb = IIf(Values.IsDark(v), foreRgb, backRgb)

            For i = 1 To moduleSize
                idx = ArrayUtil.Copy(bitmapRow, idx, colorRgb, 1, 3)
            Next
        Next

        For i = 1 To moduleSize
            Call bs.Append(bitmapRow)
        Next
    Next

    GetTrueColorBmp = Dib.GetDib(bs.Flush(), pictWidth, pictHeight, foreColorRgb, backColorRgb, False)
End Function

Public Function GetBmp(Optional ByVal moduleSize As Long = DEFAULT_MODULE_SIZE, _
                       Optional ByVal monochrome As Boolean = False, _
                       Optional ByVal foreRgb As String = ColorCode.BLACK, _
                       Optional ByVal backRgb As String = ColorCode.WHITE) As Byte()
    If m_dataBitCounter = 0 Then Call Err.Raise(51)

    If moduleSize < MIN_MODULE_SIZE Then Call Err.Raise(5)
    If ColorCode.IsWebColor(foreRgb) = False Then Call Err.Raise(5)
    If ColorCode.IsWebColor(backRgb) = False Then Call Err.Raise(5)

    Dim foreColorRgb As Long
    foreColorRgb = ColorCode.ToRGB(foreRgb)

    Dim backColorRgb As Long
    backColorRgb = ColorCode.ToRGB(backRgb)

    If monochrome Then
        GetBmp = GetMonochromeBmp(moduleSize, foreColorRgb, backColorRgb)
    Else
        GetBmp = GetTrueColorBmp(moduleSize, foreColorRgb, backColorRgb)
    End If
End Function

Public Function GetPicture(Optional ByVal moduleSize As Long = DEFAULT_MODULE_SIZE, _
                           Optional ByVal foreRgb As String = ColorCode.BLACK, _
                           Optional ByVal backRgb As String = ColorCode.WHITE, _
                           Optional ByVal picType As OlePicType = OlePicType.Bitmap) As stdole.IPictureDisp
    If m_dataBitCounter = 0 Then Call Err.Raise(51)

    If moduleSize < MIN_MODULE_SIZE Then Call Err.Raise(5)
    If ColorCode.IsWebColor(foreRgb) = False Then Call Err.Raise(5)
    If ColorCode.IsWebColor(backRgb) = False Then Call Err.Raise(5)

    Dim foreColorRgb As Long
    foreColorRgb = ColorCode.ToRGB(foreRgb)

    Dim backColorRgb As Long
    backColorRgb = ColorCode.ToRGB(backRgb)

    Dim dibData() As Byte

#If VBA7 Then
    Dim hEmf As LongPtr
#Else
    Dim hEmf As Long
#End If

    Dim ret As stdole.IPictureDisp

    Select Case picType
        Case OlePicType.Bitmap
            dibData = GetTrueColorBmp(moduleSize, foreColorRgb, backColorRgb)
            Set ret = ImageConverter.ConvertFromDib(dibData)
        Case OlePicType.EnhMetaFile
            hEmf = GetEmf(moduleSize, foreRgb)
            Set ret = ImageConverter.ConvertFromEmf(hEmf)
        Case Else
            Call Err.Raise(5)
    End Select

    Set GetPicture = ret
End Function

Private Function GetMonochromePng(ByVal moduleSize As Long, _
                                  ByVal foreColorRgb As Long, _
                                  ByVal backColorRgb As Long) As Byte()
    Dim moduleMatrix() As Variant
    moduleMatrix = GetModuleMatrix()

    Dim moduleCount As Long
    moduleCount = UBound(moduleMatrix) + 1

    Dim pictWidth As Long
    pictWidth = moduleCount * moduleSize

    Dim pictHeight As Long
    pictHeight = pictWidth

    Dim pack8bit As Long
    If pictWidth Mod 8 > 0 Then
        pack8bit = 8 - (pictWidth Mod 8)
    End If

    Dim rowSize As Long
    rowSize = 1 + (pictWidth + pack8bit) \ 8

    Dim bitmapData() As Byte
    ReDim bitmapData(rowSize * pictHeight - 1)

    Dim bs As New BitSequence

    Dim pixelColor  As Long
    Dim bitmapRow() As Byte

    Dim offset As Long
    offset = 0

    Dim filterType As Byte
    filterType = 0

    Dim r As Long
    Dim v As Variant
    Dim i As Long
    For r = 0 To UBound(moduleMatrix)
        Call bs.Clear

        Call bs.Append(filterType, 8)

        For Each v In moduleMatrix(r)
            pixelColor = IIf(Values.IsDark(v), 0, 1)

            For i = 1 To moduleSize
                Call bs.Append(pixelColor, 1)
            Next
        Next
        Call bs.Append(0, pack8bit)
        bitmapRow = bs.GetBytes()

        For i = 1 To moduleSize
            offset = ArrayUtil.CopyAll(bitmapData, offset, bitmapRow)
        Next
    Next

    GetMonochromePng = Png.GetPng( _
        bitmapData, pictWidth, pictHeight, foreColorRgb, backColorRgb, PngColorType.pIndexColor)
End Function

Private Function GetTrueColorPng(ByVal moduleSize As Long, _
                                 ByVal foreColorRgb As Long, _
                                 ByVal backColorRgb As Long, _
                                 ByVal bkStyle As BackStyle) As Byte()
    Dim moduleMatrix() As Variant
    moduleMatrix = GetModuleMatrix()

    Dim pictWidth As Long
    pictWidth = (UBound(moduleMatrix(0)) + 1) * moduleSize

    Dim pictHeight As Long
    pictHeight = pictWidth

    Dim rowSize As Long
    rowSize = 1 + IIf(bkStyle = BackStyle.bkTransparent, 4, 3) * pictWidth

    Dim offset As Long
    offset = 0

    Dim foreRgb() As Byte
    foreRgb = BitConverter.GetBytes(foreColorRgb)

    Dim backRgb() As Byte
    backRgb = BitConverter.GetBytes(backColorRgb)

    If bkStyle = BackStyle.bkTransparent Then
        foreRgb(3) = &HFF
        backRgb(3) = 0
    Else
        ReDim Preserve foreRgb(2)
        ReDim Preserve backRgb(2)
    End If

    Dim colorRgb()  As Byte
    Dim bitmapRow() As Byte
    Dim idx         As Long

    Dim filterType As Byte
    filterType = 0

    Dim bs As New ByteSequence

    Dim r As Long
    Dim v As Variant
    Dim i As Long
    For r = 0 To UBound(moduleMatrix)
        ReDim bitmapRow(rowSize - 1)
        idx = 0

        bitmapRow(idx) = filterType
        idx = idx + 1

        For Each v In moduleMatrix(r)
            colorRgb = IIf(Values.IsDark(v), foreRgb, backRgb)

            For i = 1 To moduleSize
                idx = ArrayUtil.CopyAll(bitmapRow, idx, colorRgb)
            Next
        Next

        For i = 1 To moduleSize
            Call bs.Append(bitmapRow)
        Next
    Next

    Dim tColor As PngColorType
    If bkStyle = BackStyle.bkTransparent Then
        tColor = PngColorType.pTrueColorAlpha
    Else
        tColor = PngColorType.pTrueColor
    End If

    GetTrueColorPng = Png.GetPng( _
        bs.Flush(), pictWidth, pictHeight, foreColorRgb, backColorRgb, tColor)
End Function

Public Function GetPng(Optional ByVal moduleSize As Long = DEFAULT_MODULE_SIZE, _
                       Optional ByVal monochrome As Boolean = False, _
                       Optional ByVal foreRgb As String = ColorCode.BLACK, _
                       Optional ByVal backRgb As String = ColorCode.WHITE, _
                       Optional ByVal bkStyle As BackStyle = BackStyle.bkOpaque) As Byte()
    If m_dataBitCounter = 0 Then Call Err.Raise(51)

    If moduleSize < MIN_MODULE_SIZE Then Call Err.Raise(5)
    If ColorCode.IsWebColor(foreRgb) = False Then Call Err.Raise(5)
    If ColorCode.IsWebColor(backRgb) = False Then Call Err.Raise(5)

    Dim foreColorRgb As Long
    foreColorRgb = ColorCode.ToRGB(foreRgb)

    Dim backColorRgb As Long
    backColorRgb = ColorCode.ToRGB(backRgb)

    If monochrome Then
        GetPng = GetMonochromePng(moduleSize, foreColorRgb, backColorRgb)
    Else
        GetPng = GetTrueColorPng(moduleSize, foreColorRgb, backColorRgb, bkStyle)
    End If
End Function

Public Function GetSvg(Optional ByVal moduleSize As Long = DEFAULT_MODULE_SIZE, _
                       Optional ByVal foreRgb As String = ColorCode.BLACK) As String
    If m_dataBitCounter = 0 Then Call Err.Raise(51)

    If moduleSize < MIN_MODULE_SIZE Then Call Err.Raise(5)
    If ColorCode.IsWebColor(foreRgb) = False Then Call Err.Raise(5)

    Dim foreColorRgb As Long
    foreColorRgb = ColorCode.ToRGB(foreRgb)

    Dim moduleMatrix() As Variant
    moduleMatrix = GetModuleMatrix()

    Dim pictWidth As Long
    pictWidth = (UBound(moduleMatrix(0)) + 1) * moduleSize

    Dim pictHeight As Long
    pictHeight = pictWidth

    Dim pixelData() As Variant
    ReDim pixelData(pictHeight - 1)

    Dim imgRow() As Long

    Dim r As Long
    Dim c As Long

    Dim rowArray As Variant
    Dim v As Variant
    Dim clr As Long
    Dim i As Long
    For Each rowArray In moduleMatrix
        ReDim imgRow(pictWidth - 1)
        c = 0

        For Each v In rowArray
            clr = IIf(Values.IsDark(v), 1, 0)
            For i = 1 To moduleSize
                imgRow(c) = clr
                c = c + 1
            Next
        Next

        For i = 1 To moduleSize
            pixelData(r) = imgRow
            r = r + 1
        Next
    Next

    GetSvg = Svg.GetSvg(pixelData, pictWidth, pictHeight, foreColorRgb)
End Function

#If VBA7 Then
Public Function GetEmf(Optional ByVal moduleSize As Long = DEFAULT_MODULE_SIZE, _
                       Optional ByVal foreRgb As String = ColorCode.BLACK) As LongPtr
#Else
Public Function GetEmf(Optional ByVal moduleSize As Long = DEFAULT_MODULE_SIZE, _
                       Optional ByVal foreRgb As String = ColorCode.BLACK) As Long
#End If
    If m_dataBitCounter = 0 Then Call Err.Raise(51)

    If moduleSize < MIN_MODULE_SIZE Then Call Err.Raise(5)
    If ColorCode.IsWebColor(foreRgb) = False Then Call Err.Raise(5)

    Dim foreColorRgb As Long
    foreColorRgb = ColorCode.ToRGB(foreRgb)

    Dim moduleMatrix() As Variant
    moduleMatrix = GetModuleMatrix()

    Dim pictWidth As Long
    pictWidth = (UBound(moduleMatrix(0)) + 1) * moduleSize

    Dim pictHeight As Long
    pictHeight = pictWidth

    Dim pixelData() As Variant
    ReDim pixelData(pictHeight - 1)

    Dim imgRow() As Long

    Dim r As Long
    Dim c As Long

    Dim rowArray As Variant
    Dim v As Variant
    Dim clr As Long
    Dim i As Long
    For Each rowArray In moduleMatrix
        ReDim imgRow(pictWidth - 1)
        c = 0

        For Each v In rowArray
            clr = IIf(Values.IsDark(v), 1, 0)
            For i = 1 To moduleSize
                imgRow(c) = clr
                c = c + 1
            Next
        Next

        For i = 1 To moduleSize
            pixelData(r) = imgRow
            r = r + 1
        Next
    Next

    GetEmf = Emf.GetEmf(pixelData, pictWidth, pictHeight, foreColorRgb)
End Function

Private Function GetBilevelTiff(ByVal moduleSize As Long) As Byte()
    Dim moduleMatrix() As Variant
    moduleMatrix = GetModuleMatrix()

    Dim moduleCount As Long
    moduleCount = UBound(moduleMatrix) + 1

    Dim pictWidth As Long
    pictWidth = moduleCount * moduleSize

    Dim pictHeight As Long
    pictHeight = pictWidth

    Dim pack8bit As Long
    If pictWidth Mod 8 > 0 Then
        pack8bit = 8 - (pictWidth Mod 8)
    End If

    Dim rowSize As Long
    rowSize = (pictWidth + pack8bit) \ 8

    Dim bitmapData() As Byte
    ReDim bitmapData(rowSize * pictHeight - 1)

    Dim bs As New BitSequence

    Dim pixelColor  As Long
    Dim bitmapRow() As Byte

    Dim offset As Long
    offset = 0

    Dim r As Long
    Dim v As Variant
    Dim i As Long
    For r = 0 To UBound(moduleMatrix)
        Call bs.Clear

        For Each v In moduleMatrix(r)
            pixelColor = IIf(Values.IsDark(v), 0, 1)

            For i = 1 To moduleSize
                Call bs.Append(pixelColor, 1)
            Next
        Next
        Call bs.Append(0, pack8bit)
        bitmapRow = bs.GetBytes()

        For i = 1 To moduleSize
            offset = ArrayUtil.CopyAll(bitmapData, offset, bitmapRow)
        Next
    Next

    GetBilevelTiff = Tiff.GetTiff( _
        bitmapData, pictWidth, pictHeight, TiffImageType.Bilevel, Empty)
End Function

Private Function GetPaletteColorTiff(ByVal moduleSize As Long, _
                                     ByVal foreColorRgb As Long, _
                                     ByVal backColorRgb As Long) As Byte()
    Dim moduleMatrix() As Variant
    moduleMatrix = GetModuleMatrix()

    Dim pictWidth As Long
    pictWidth = (UBound(moduleMatrix(0)) + 1) * moduleSize

    Dim pictHeight As Long
    pictHeight = pictWidth

    Dim rowSize As Long
    rowSize = pictWidth

    Dim bitmapData() As Byte
    ReDim bitmapData(rowSize * pictHeight - 1)

    Dim offset As Long
    offset = 0

    Dim paletteIndex As Long
    Dim bitmapRow()  As Byte
    Dim idx          As Long

    Dim bs As New ByteSequence

    Dim r As Long
    Dim v As Variant
    Dim i As Long
    For r = 0 To UBound(moduleMatrix)
        ReDim bitmapRow(rowSize - 1)
        idx = 0

        For Each v In moduleMatrix(r)
            paletteIndex = IIf(Values.IsDark(v), 0, 1)

            For i = 1 To moduleSize
                bitmapRow(idx) = CByte(paletteIndex)
                idx = idx + 1
            Next
        Next

        For i = 1 To moduleSize
            Call bs.Append(bitmapRow)
        Next
    Next

    GetPaletteColorTiff = Tiff.GetTiff( _
        bs.Flush(), pictWidth, pictHeight, TiffImageType.PaletteColor, Array(foreColorRgb, backColorRgb))
End Function

Private Function GetFullColorTiff(ByVal moduleSize As Long, _
                                  ByVal foreColorRgb As Long, _
                                  ByVal backColorRgb As Long) As Byte()
    Dim moduleMatrix() As Variant
    moduleMatrix = GetModuleMatrix()

    Dim pictWidth As Long
    pictWidth = (UBound(moduleMatrix(0)) + 1) * moduleSize

    Dim pictHeight As Long
    pictHeight = pictWidth

    Dim rowSize As Long
    rowSize = 3 * pictWidth

    Dim foreRgb() As Byte
    foreRgb = BitConverter.GetBytes(foreColorRgb)
    ReDim Preserve foreRgb(2)

    Dim backRgb() As Byte
    backRgb = BitConverter.GetBytes(backColorRgb)
    ReDim Preserve backRgb(2)

    Dim colorRgb()  As Byte
    Dim bitmapRow() As Byte
    Dim idx         As Long

    Dim bs As New ByteSequence

    Dim r As Long
    Dim v As Variant
    Dim i As Long
    For r = 0 To UBound(moduleMatrix)
        ReDim bitmapRow(rowSize - 1)
        idx = 0

        For Each v In moduleMatrix(r)
            colorRgb = IIf(Values.IsDark(v), foreRgb, backRgb)

            For i = 1 To moduleSize
                idx = ArrayUtil.CopyAll(bitmapRow, idx, colorRgb)
            Next
        Next

        For i = 1 To moduleSize
            Call bs.Append(bitmapRow)
        Next
    Next

    GetFullColorTiff = Tiff.GetTiff( _
        bs.Flush(), pictWidth, pictHeight, TiffImageType.FullColor, Empty)
End Function

Public Function GetTiff(Optional ByVal moduleSize As Long = DEFAULT_MODULE_SIZE, _
                        Optional ByVal monochrome As Boolean = False, _
                        Optional ByVal bilevelImage As Boolean = False, _
                        Optional ByVal foreRgb As String = ColorCode.BLACK, _
                        Optional ByVal backRgb As String = ColorCode.WHITE) As Byte()
    If m_dataBitCounter = 0 Then Call Err.Raise(51)

    If moduleSize < MIN_MODULE_SIZE Then Call Err.Raise(5)
    If ColorCode.IsWebColor(foreRgb) = False Then Call Err.Raise(5)
    If ColorCode.IsWebColor(backRgb) = False Then Call Err.Raise(5)

    Dim foreColorRgb As Long
    foreColorRgb = ColorCode.ToRGB(foreRgb)

    Dim backColorRgb As Long
    backColorRgb = ColorCode.ToRGB(backRgb)

    If bilevelImage Then
        GetTiff = GetBilevelTiff(moduleSize)
    ElseIf monochrome Then
        GetTiff = GetPaletteColorTiff(moduleSize, foreColorRgb, backColorRgb)
    Else
        GetTiff = GetFullColorTiff(moduleSize, foreColorRgb, backColorRgb)
    End If
End Function

Public Function GetGif(Optional ByVal moduleSize As Long = DEFAULT_MODULE_SIZE, _
                       Optional ByVal foreRgb As String = ColorCode.BLACK, _
                       Optional ByVal backRgb As String = ColorCode.WHITE, _
                       Optional ByVal bkStyle As BackStyle = BackStyle.bkOpaque, _
                       Optional ByVal bpp As Long = 1) As Byte()
    If m_dataBitCounter = 0 Then Call Err.Raise(51)

    If ColorCode.IsWebColor(foreRgb) = False Then Call Err.Raise(5)
    If ColorCode.IsWebColor(backRgb) = False Then Call Err.Raise(5)
    If (1 <= bpp And bpp <= 8) = False Then Call Err.Raise(5)

    Dim foreColorRgb As Long
    foreColorRgb = ColorCode.ToRGB(foreRgb)

    Dim backColorRgb As Long
    backColorRgb = ColorCode.ToRGB(backRgb)

    Dim moduleMatrix() As Variant
    moduleMatrix = GetModuleMatrix()

    Dim pictWidth As Long
    pictWidth = (UBound(moduleMatrix(0)) + 1) * moduleSize

    Dim pictHeight As Long
    pictHeight = pictWidth

    Dim paletteIndex As Long
    Dim bitmapRow()  As Byte
    Dim idx          As Long

    Dim bs As New ByteSequence

    Dim r As Long
    Dim v As Variant
    Dim i As Long
    For r = 0 To UBound(moduleMatrix)
        ReDim bitmapRow(pictWidth - 1)
        idx = 0

        For Each v In moduleMatrix(r)
            paletteIndex = IIf(Values.IsDark(v), 0, 1)

            For i = 1 To moduleSize
                bitmapRow(idx) = CByte(paletteIndex)
                idx = idx + 1
            Next
        Next

        For i = 1 To moduleSize
            Call bs.Append(bitmapRow)
        Next
    Next

    Dim palette() As Long
    ReDim palette(2 ^ bpp - 1)

    palette(0) = foreColorRgb
    palette(1) = backColorRgb

    Dim transparentIndex As Long
    transparentIndex = 1

    Dim bkTransparent As Boolean
    bkTransparent = (bkStyle = BackStyle.bkTransparent)

    GetGif = Gif.GetGif(bs.Flush(), pictWidth, pictHeight, palette, bkTransparent, transparentIndex)
End Function

Public Sub SaveAs(ByVal filePath As String, _
                  Optional ByVal moduleSize As Long = DEFAULT_MODULE_SIZE, _
                  Optional ByVal foreRgb As String = ColorCode.BLACK, _
                  Optional ByVal backRgb As String = ColorCode.WHITE, _
                  Optional ByVal fmt As ImageFormat = ImageFormat.fmtBmp + ImageFormat.fmtMonochrome, _
                  Optional ByVal bkStyle As BackStyle = BackStyle.bkOpaque)
    If m_dataBitCounter = 0 Then Call Err.Raise(51)

    If Len(filePath) = 0 Then Call Err.Raise(5)
    If moduleSize < MIN_MODULE_SIZE Then Call Err.Raise(5)
    If ColorCode.IsWebColor(foreRgb) = False Then Call Err.Raise(5)
    If ColorCode.IsWebColor(backRgb) = False Then Call Err.Raise(5)

    Dim binData() As Byte
    Dim strData   As String
    Dim pict      As stdole.IPictureDisp

    If (fmt And &HF0&) = 0 Then
        fmt = fmt + ImageFormat.fmtBmp
    End If

    Select Case fmt And &HF0&
        Case ImageFormat.fmtBmp
            Select Case fmt And &HF
                Case ImageFormat.fmtMonochrome
                    binData = GetBmp(moduleSize, True, foreRgb, backRgb)
                Case ImageFormat.fmtTrueColor
                    binData = GetBmp(moduleSize, False, foreRgb, backRgb)
                Case Else
                    Call Err.Raise(5)
            End Select

        Case ImageFormat.fmtEmf
            Set pict = GetPicture(moduleSize, foreRgb, backRgb, picType:=OlePicType.EnhMetaFile)

        Case ImageFormat.fmtGif
            binData = GetGif(moduleSize, foreRgb, backRgb, bkStyle, 1)

        Case ImageFormat.fmtPng
            Select Case fmt And &HF
                Case ImageFormat.fmtMonochrome
                    binData = GetPng(moduleSize, True, foreRgb, backRgb)
                Case ImageFormat.fmtTrueColor
                    binData = GetPng(moduleSize, False, foreRgb, backRgb, bkStyle)
                Case Else
                    Call Err.Raise(5)
            End Select

        Case ImageFormat.fmtSvg
            strData = GetSvg(moduleSize, foreRgb)

        Case ImageFormat.fmtTiff
            Select Case fmt And &HF
                Case ImageFormat.fmtMonochrome
                    binData = GetTiff(moduleSize, True, False, foreRgb, backRgb)
                Case ImageFormat.fmtTrueColor
                    binData = GetTiff(moduleSize, False, False, foreRgb, backRgb)
                Case ImageFormat.fmtBilevel
                    binData = GetTiff(moduleSize, True, True)
                Case Else
                    Call Err.Raise(5)
            End Select

        Case Else
            Call Err.Raise(5)
    End Select

    Dim fNum As Integer

    Select Case fmt And &HF0&
        Case ImageFormat.fmtBmp, ImageFormat.fmtPng, ImageFormat.fmtTiff, ImageFormat.fmtGif
            fNum = FreeFile()
            Open filePath For Binary As #fNum
            Put #fNum, , binData
            Close #fNum
        Case ImageFormat.fmtSvg
            fNum = FreeFile()
            Open filePath For Output As #fNum Len = &H7FFF
            Print #fNum, strData
            Close #fNum
        Case ImageFormat.fmtEmf
            Call stdole.SavePicture(pict, filePath)
        Case Else
            Call Err.Raise(5)
    End Select
End Sub

Public Sub SetToClipboard(Optional ByVal moduleSize As Long = DEFAULT_MODULE_SIZE, _
                          Optional ByVal foreRgb As String = ColorCode.BLACK, _
                          Optional ByVal backRgb As String = ColorCode.WHITE, _
                          Optional ByVal fmt As ImageFormat = ImageFormat.fmtBmp)
    If m_dataBitCounter = 0 Then Call Err.Raise(51)

    If moduleSize < MIN_MODULE_SIZE Then Call Err.Raise(5)
    If ColorCode.IsWebColor(foreRgb) = False Then Call Err.Raise(5)
    If ColorCode.IsWebColor(backRgb) = False Then Call Err.Raise(5)

    Dim bytes() As Byte

#If VBA7 Then
    Dim hEmf As LongPtr
#Else
    Dim hEmf As Long
#End If

    Select Case fmt And &HF0&
        Case ImageFormat.fmtBmp
            bytes = GetBmp(moduleSize, False, foreRgb, backRgb)
            Call ClipboardUtil.SetDib(bytes)
        Case ImageFormat.fmtEmf
            hEmf = GetEmf(moduleSize, foreRgb)
            Call ClipboardUtil.SetEmf(hEmf)
        Case Else
            Call Err.Raise(5)
    End Select
End Sub
