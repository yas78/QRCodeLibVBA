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

Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByVal pDest As Long, ByVal pSrc As Long, ByVal sz As Long)

Private m_parent As Symbols

Private m_position As Long

Private m_currEncoder       As IQRCodeEncoder
Private m_currEncodingMode  As EncodingMode
Private m_currVersion       As Long

Private m_dataBitCapacity   As Long
Private m_dataBitCounter    As Long

Private m_segments          As New Collection ' Item : IQRCodeEncoder
Private m_segmentCounter    As New Dictionary ' Key : EncodingMode, Item : Long

'------------------------------------------------------------------------------
' (概要)
'  オブジェクトを初期化します。
'------------------------------------------------------------------------------
Friend Sub Initialize(ByVal parentObj As Symbols)
    
    Set m_parent = parentObj

    m_position = parentObj.Count
    
    Set m_currEncoder = Nothing
    m_currEncodingMode = EncodingMode.UNKNOWN
    m_currVersion = parentObj.MinVersion
    
    m_dataBitCapacity = 8 * DataCodeword.GetTotalNumber( _
        parentObj.ErrorCorrectionLevel, parentObj.MinVersion)
        
    m_dataBitCounter = 0
    
    Call m_segmentCounter.Add(EncodingMode.NUMERIC, 0)
    Call m_segmentCounter.Add(EncodingMode.ALPHA_NUMERIC, 0)
    Call m_segmentCounter.Add(EncodingMode.EIGHT_BIT_BYTE, 0)
    Call m_segmentCounter.Add(EncodingMode.KANJI, 0)

    If parentObj.StructuredAppendAllowed Then
        m_dataBitCapacity = m_dataBitCapacity - StructuredAppend.HEADER_LENGTH
    End If
    
End Sub

'------------------------------------------------------------------------------
' (概要)
'  親オブジェクトを取得します。
'------------------------------------------------------------------------------
Public Property Get Parent() As Symbols

    Set Parent = m_parent
    
End Property

'------------------------------------------------------------------------------
' (概要)
'  型番を取得します。
'------------------------------------------------------------------------------
Public Property Get Version() As Long
    
    Version = m_currVersion

End Property

'------------------------------------------------------------------------------
' (概要)
'  誤り訂正レベルを取得します。
'------------------------------------------------------------------------------
Friend Property Get CurrentEncodingMode() As EncodingMode

    CurrentEncodingMode = m_currEncodingMode
    
End Property

'------------------------------------------------------------------------------
' (概要)
'  シンボルに文字を追加します。
'
' (戻り値)
'  正常終了時は True を返します。シンボル容量が不足している場合は False を返します。
'------------------------------------------------------------------------------
Friend Function TryAppend(ByVal c As String) As Boolean

#If [DEBUG] Then
    Debug.Assert Len(c) = 1
#End If

    Dim bitLength  As Long
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

'------------------------------------------------------------------------------
' (概要)
'  符号化モードを設定します。
'
' (戻り値)
'  正常終了は True、シンボル容量が不足している場合は False を返します。
'------------------------------------------------------------------------------
Friend Function TrySetEncodingMode( _
    ByVal encMode As EncodingMode, ByVal c As String) As Boolean

#If [DEBUG] Then
    Debug.Assert Len(c) = 1
#End If

    Dim encoder As IQRCodeEncoder
    Set encoder = QRCodeEncoderFactory.CreateEncoder( _
            encMode, m_parent.ByteModeCharset)

    Dim bitLength As Long
    bitLength = encoder.GetCodewordBitLength(c)

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
    
    Set m_currEncoder = encoder
    Call m_segments.Add(encoder)
    m_segmentCounter(encMode) = m_segmentCounter(encMode) + 1
    m_currEncodingMode = encMode
    
    TrySetEncodingMode = True

End Function

'------------------------------------------------------------------------------
' (概要)
'  型番を決定します。
'------------------------------------------------------------------------------
Private Sub SelectVersion()
    
    Dim encMode As Variant ' Subtype : EncodingMode
    
    For Each encMode In m_segmentCounter.Keys
        Dim num As Long
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

    If m_parent.StructuredAppendAllowed Then
        m_dataBitCapacity = m_dataBitCapacity - StructuredAppend.HEADER_LENGTH
    End If

End Sub

'------------------------------------------------------------------------------
' (概要)
'  データ領域ブロックを、内部形式 Byte型配列のVariant型配列として返します。
'------------------------------------------------------------------------------
Private Function BuildDataBlock() As Variant() ' Subtype : Byte()

    Dim dataBytes() As Byte
    dataBytes = GetMessageBytes()
    
    Dim numPreBlocks As Long
    numPreBlocks = RSBlock.GetTotalNumber( _
        m_parent.ErrorCorrectionLevel, m_currVersion, True)

    Dim numFolBlocks As Long
    numFolBlocks = RSBlock.GetTotalNumber( _
        m_parent.ErrorCorrectionLevel, m_currVersion, False)
    
    Dim ret() As Variant ' Subtype : Byte()
    ReDim ret(numPreBlocks + numFolBlocks - 1)
    
    Dim dataIdx As Long
    dataIdx = 0

    Dim numPreBlockDataCodewords As Long
    numPreBlockDataCodewords = RSBlock.GetNumberDataCodewords( _
        m_parent.ErrorCorrectionLevel, m_currVersion, True)
    
    Dim data() As Byte
    Dim i As Long
    Dim j As Long
    
    For i = 0 To numPreBlocks - 1
        ReDim data(numPreBlockDataCodewords - 1)
        
        For j = 0 To UBound(data)
            data(j) = dataBytes(dataIdx)
            dataIdx = dataIdx + 1
        Next
        
        ret(i) = data
    Next

    Dim numFolBlockDataCodewords As Long
    numFolBlockDataCodewords = RSBlock.GetNumberDataCodewords( _
        m_parent.ErrorCorrectionLevel, m_currVersion, False)
    
    For i = numPreBlocks To numPreBlocks + numFolBlocks - 1
        ReDim data(numFolBlockDataCodewords - 1)
        
        For j = 0 To UBound(data)
            data(j) = dataBytes(dataIdx)
            dataIdx = dataIdx + 1
        Next
        
        ret(i) = data
    Next
    
    BuildDataBlock = ret
    
End Function

'------------------------------------------------------------------------------
' (概要)
'  誤り訂正領域ブロックを、内部形式 Byte型配列のVariant型配列として返します。
'------------------------------------------------------------------------------
Private Function BuildErrorCorrectionBlock( _
    ByRef dataBlock() As Variant) As Variant() ' Subtype : Byte()

    Dim i As Long
    Dim j As Long

    Dim numECCodewords As Long
    numECCodewords = RSBlock.GetNumberECCodewords( _
        m_parent.ErrorCorrectionLevel, m_currVersion)
    
    Dim numPreBlocks As Long
    numPreBlocks = RSBlock.GetTotalNumber( _
        m_parent.ErrorCorrectionLevel, m_currVersion, True)
    
    Dim numFolBlocks As Long
    numFolBlocks = RSBlock.GetTotalNumber( _
        m_parent.ErrorCorrectionLevel, m_currVersion, False)
    
    Dim ret() As Variant ' Subtype : Byte()
    ReDim ret(numPreBlocks + numFolBlocks - 1)
    
    Dim eccDataTmp() As Byte
    ReDim eccDataTmp(numECCodewords - 1)
    
    For i = 0 To UBound(ret)
        ret(i) = eccDataTmp
    Next

    Dim gp() As Variant
    gp = GeneratorPolynomials.Item(numECCodewords)
    
    Dim eccIdx      As Long
    Dim blockIdx    As Long
    Dim data()      As Long
    Dim exp         As Long
    
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

'------------------------------------------------------------------------------
' (概要)
'  符号化領域のバイトデータを返します。
'------------------------------------------------------------------------------
Private Function GetEncodingRegionBytes() As Byte()

    Dim dataBlock() As Variant ' Subtype : Byte()
    dataBlock = BuildDataBlock()
    
    Dim ecBlock() As Variant ' Subtype : Byte()
    ecBlock = BuildErrorCorrectionBlock(dataBlock)
    
    Dim numCodewords  As Long
    numCodewords = Codeword.GetTotalNumber(m_currVersion)
    
    Dim numDataCodewords As Long
    numDataCodewords = DataCodeword.GetTotalNumber( _
        m_parent.ErrorCorrectionLevel, m_currVersion)
    
    Dim ret() As Byte
    ReDim ret(numCodewords - 1)
    
    Dim r As Long
    Dim c As Long
    
    Dim idx As Long
    idx = 0

    Dim n As Long
    n = 0
    
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

'------------------------------------------------------------------------------
' (概要)
'  コード語に変換するメッセージビット列を返します。
'------------------------------------------------------------------------------
Private Function GetMessageBytes() As Byte()
    
    Dim bs As BitSequence
    Set bs = New BitSequence

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

#If [DEBUG] Then
    Debug.Assert m_parent.Count > 1
#End If

    Call bs.Append(ModeIndicator.STRUCTURED_APPEND_VALUE, _
                   ModeIndicator.Length)
    Call bs.Append(m_position, _
                   SymbolSequenceIndicator.POSITION_LENGTH)
    Call bs.Append(m_parent.Count - 1, _
                   SymbolSequenceIndicator.TOTAL_NUMBER_LENGTH)
    Call bs.Append(m_parent.StructuredAppendParity, _
                   StructuredAppend.PARITY_DATA_LENGTH)
    
End Sub

Private Sub WriteSegments(ByVal bs As BitSequence)
    
    Dim i As Long
    Dim data() As Byte
    Dim codewordBitLength As Long
    
    Dim segment As IQRCodeEncoder
    
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

    Dim flag As Boolean
    flag = True
    
    Do While bs.Length < 8 * numDataCodewords
        Call bs.Append(IIf(flag, 236, 17), 8)
        flag = Not flag
    Loop

End Sub

'------------------------------------------------------------------------------
' (概要)
'  シンボルの明暗パターンを返します。
'------------------------------------------------------------------------------
Private Function GetModuleMatrix() As Variant()

    Dim numModulesPerSide As Long
    numModulesPerSide = Module.GetNumModulesPerSide(m_currVersion)
    
    Dim moduleMatrix() As Variant
    ReDim moduleMatrix(numModulesPerSide - 1)
    
    Dim i As Long
    Dim cols() As Long
    
    For i = 0 To UBound(moduleMatrix)
        ReDim cols(numModulesPerSide - 1)
        moduleMatrix(i) = cols
    Next
    
    Call FinderPattern.Place(moduleMatrix)
    Call Separator.Place(moduleMatrix)
    Call TimingPattern.Place(moduleMatrix)
    
    If m_currVersion >= 2 Then
        Call AlignmentPattern.Place(moduleMatrix, m_currVersion)
    End If
    
    Call FormatInfo.PlaceTempBlank(moduleMatrix)
    
    If m_currVersion >= 7 Then
        Call VersionInfo.PlaceTempBlank(moduleMatrix)
    End If
    
    Call PlaceSymbolChar(moduleMatrix)
    Call RemainderBit.Place(moduleMatrix)

    Dim maskPatternReference As Long
    maskPatternReference = Masking.Apply( _
        moduleMatrix, m_currVersion, m_parent.ErrorCorrectionLevel)

    Call FormatInfo.Place(moduleMatrix, _
                          m_parent.ErrorCorrectionLevel, _
                          maskPatternReference)
    
    If m_currVersion >= 7 Then
        Call VersionInfo.Place(moduleMatrix, m_currVersion)
    End If

    GetModuleMatrix = moduleMatrix
        
End Function

'------------------------------------------------------------------------------
' (概要)
' シンボルキャラクタを配置します。
'------------------------------------------------------------------------------
Private Sub PlaceSymbolChar(ByRef moduleMatrix() As Variant)

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
    
    Dim bitPos As Long
    Dim i As Long

    For i = 0 To UBound(data)
        bitPos = 7
    
        Do While bitPos >= 0
            If moduleMatrix(r)(c) = 0 Then
                moduleMatrix(r)(c) = IIf((data(i) And 2 ^ bitPos) > 0, 1, -1)
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

'------------------------------------------------------------------------------
' (概要)
'  色深度1bppビットマップファイルのバイトデータを返します。
'------------------------------------------------------------------------------
Public Function Get1bppDIB(Optional ByVal moduleSize As Long = 5, _
                           Optional ByVal foreRGB As String = "#000000", _
                           Optional ByVal backRGB As String = "#FFFFFF") As Byte()
    
    If moduleSize < 1 Then Err.Raise 5
    
    Dim foreColorRGB As Long
    foreColorRGB = ColorCode.ToRGB(foreRGB)
    Dim backColorRGB As Long
    backColorRGB = ColorCode.ToRGB(backRGB)
    
    If m_dataBitCounter = 0 Then Err.Raise 51

    Dim moduleMatrix() As Variant
    moduleMatrix = QuietZone.Place(GetModuleMatrix())
    
    Dim moduleCount As Long
    moduleCount = UBound(moduleMatrix) + 1
    
    Dim pictWidth As Long
    pictWidth = moduleCount * moduleSize
    
    Dim pictHeight As Long
    pictHeight = moduleCount * moduleSize
    
    Dim hByteLen As Long
    hByteLen = (pictWidth + 7) \ 8
    
    Dim pack8bit As Long
    If pictWidth Mod 8 > 0 Then
        pack8bit = 8 - (pictWidth Mod 8)
    End If
    
    Dim pack32bit As Long
    If hByteLen Mod 4 > 0 Then
        pack32bit = 8 * (4 - (hByteLen Mod 4))
    End If
    
    Dim bs As BitSequence
    Set bs = New BitSequence
    
    Dim r As Long
    Dim c As Long
    Dim i As Long
    Dim j As Long
    
    For r = UBound(moduleMatrix) To 0 Step -1
        For i = 1 To moduleSize
            For c = 0 To UBound(moduleMatrix(r))
                For j = 1 To moduleSize
                    Call bs.Append(IIf(moduleMatrix(r)(c) > 0, 0, 1), 1)
                Next
            Next

            Call bs.Append(0, pack8bit)
            Call bs.Append(0, pack32bit)
        Next
    Next

    Dim dataBlock() As Byte
    dataBlock = bs.GetBytes()

    Dim bfh         As BITMAPFILEHEADER
    Dim bih         As BITMAPINFOHEADER
    Dim palette(1)  As RGBQUAD

    With bfh
        .bfType = &H4D42
        .bfSize = 62 + (UBound(dataBlock) + 1)
        .bfReserved1 = 0
        .bfReserved2 = 0
        .bfOffBits = 62
    End With
                
    With bih
        .biSize = 40
        .biWidth = pictWidth
        .biHeight = pictHeight
        .biPlanes = 1
        .biBitCount = 1
        .biCompression = 0
        .biSizeImage = 0
        .biXPelsPerMeter = 3780 ' 96dpi
        .biYPelsPerMeter = 3780 ' 96dpi
        .biClrUsed = 0
        .biClrImportant = 0
    End With
    
    With palette(0)
        .rgbBlue = CByte((foreColorRGB And &HFF0000) \ 2 ^ 16)
        .rgbGreen = CByte((foreColorRGB And &HFF00&) \ 2 ^ 8)
        .rgbRed = CByte(foreColorRGB And &HFF&)
        .rgbReserved = 0
    End With
    
    With palette(1)
        .rgbBlue = CByte((backColorRGB And &HFF0000) \ 2 ^ 16)
        .rgbGreen = CByte((backColorRGB And &HFF00&) \ 2 ^ 8)
        .rgbRed = CByte(backColorRGB And &HFF&)
        .rgbReserved = 0
    End With
    
    Dim ret() As Byte
    ReDim ret(62 + UBound(dataBlock))
    
    With bfh
        Call MoveMemory(VarPtr(ret(0)), VarPtr(.bfType), 2)
        Call MoveMemory(VarPtr(ret(2)), VarPtr(.bfSize), 4)
        Call MoveMemory(VarPtr(ret(6)), VarPtr(.bfReserved1), 2)
        Call MoveMemory(VarPtr(ret(8)), VarPtr(.bfReserved2), 2)
        Call MoveMemory(VarPtr(ret(10)), VarPtr(.bfOffBits), 4)
    End With

    With bih
        Call MoveMemory(VarPtr(ret(14)), VarPtr(.biSize), 4)
        Call MoveMemory(VarPtr(ret(18)), VarPtr(.biWidth), 4)
        Call MoveMemory(VarPtr(ret(22)), VarPtr(.biHeight), 4)
        Call MoveMemory(VarPtr(ret(26)), VarPtr(.biPlanes), 2)
        Call MoveMemory(VarPtr(ret(28)), VarPtr(.biBitCount), 2)
        Call MoveMemory(VarPtr(ret(30)), VarPtr(.biCompression), 4)
        Call MoveMemory(VarPtr(ret(34)), VarPtr(.biSizeImage), 4)
        Call MoveMemory(VarPtr(ret(38)), VarPtr(.biXPelsPerMeter), 4)
        Call MoveMemory(VarPtr(ret(42)), VarPtr(.biYPelsPerMeter), 4)
        Call MoveMemory(VarPtr(ret(46)), VarPtr(.biClrUsed), 4)
        Call MoveMemory(VarPtr(ret(50)), VarPtr(.biClrImportant), 4)
    End With

    With palette(0)
        Call MoveMemory(VarPtr(ret(54)), VarPtr(.rgbBlue), 1)
        Call MoveMemory(VarPtr(ret(55)), VarPtr(.rgbGreen), 1)
        Call MoveMemory(VarPtr(ret(56)), VarPtr(.rgbRed), 1)
        Call MoveMemory(VarPtr(ret(57)), VarPtr(.rgbReserved), 1)
    End With

    With palette(1)
        Call MoveMemory(VarPtr(ret(58)), VarPtr(.rgbBlue), 1)
        Call MoveMemory(VarPtr(ret(59)), VarPtr(.rgbGreen), 1)
        Call MoveMemory(VarPtr(ret(60)), VarPtr(.rgbRed), 1)
        Call MoveMemory(VarPtr(ret(61)), VarPtr(.rgbReserved), 1)
    End With
    
    Call MoveMemory(VarPtr(ret(62)), VarPtr(dataBlock(0)), UBound(dataBlock) + 1)
    
    Get1bppDIB = ret
    
End Function

'------------------------------------------------------------------------------
' (概要)
'  色深度24bppビットマップファイルのバイトデータを返します。
'------------------------------------------------------------------------------
Public Function Get24bppDIB(Optional ByVal moduleSize As Long = 5, _
                            Optional ByVal foreRGB As String = "#000000", _
                            Optional ByVal backRGB As String = "#FFFFFF") As Byte()

    If moduleSize < 1 Then Err.Raise 5
    
    Dim foreColorRGB As Long
    foreColorRGB = ColorCode.ToRGB(foreRGB)
    Dim backColorRGB As Long
    backColorRGB = ColorCode.ToRGB(backRGB)
    
    If m_dataBitCounter = 0 Then Err.Raise 51

    Dim moduleMatrix() As Variant
    moduleMatrix = QuietZone.Place(GetModuleMatrix())

    Dim pictWidth As Long
    pictWidth = (UBound(moduleMatrix) + 1) * moduleSize
    
    Dim pictHeight As Long
    pictHeight = pictWidth

    Dim hByteLen As Long
    hByteLen = 3 * pictWidth

    Dim pack4byte As Long
    If hByteLen Mod 4 > 0 Then
        pack4byte = 4 - (hByteLen Mod 4)
    End If

    Dim dataBlock() As Byte
    ReDim dataBlock((hByteLen + pack4byte) * pictHeight - 1)

    Dim r As Long
    Dim c As Long
    Dim i As Long
    Dim j As Long
    
    Dim colorRGB As Long
    
    Dim idx As Long
    idx = 0
    
    For r = UBound(moduleMatrix) To 0 Step -1
        For i = 1 To moduleSize
            For c = 0 To UBound(moduleMatrix(r))
                For j = 1 To moduleSize
                    colorRGB = IIf(moduleMatrix(r)(c) > 0, foreColorRGB, backColorRGB)
                    dataBlock(idx + 0) = CByte((colorRGB And &HFF0000) \ 2 ^ 16) ' B
                    dataBlock(idx + 1) = CByte((colorRGB And &HFF00&) \ 2 ^ 8)   ' G
                    dataBlock(idx + 2) = CByte(colorRGB And &HFF&)               ' R
                    idx = idx + 3
                Next
            Next

            idx = idx + pack4byte
        Next
    Next

    Dim bfh As BITMAPFILEHEADER
    Dim bih As BITMAPINFOHEADER

    With bfh
        .bfType = &H4D42
        .bfSize = 54 + (UBound(dataBlock) + 1)
        .bfReserved1 = 0
        .bfReserved2 = 0
        .bfOffBits = 54
    End With

    With bih
        .biSize = 40
        .biWidth = pictWidth
        .biHeight = pictHeight
        .biPlanes = 1
        .biBitCount = 24
        .biCompression = 0
        .biSizeImage = 0
        .biXPelsPerMeter = 3780 ' 96dpi
        .biYPelsPerMeter = 3780 ' 96dpi
        .biClrUsed = 0
        .biClrImportant = 0
    End With

    Dim ret() As Byte
    ReDim ret(54 + UBound(dataBlock))
    
    With bfh
        Call MoveMemory(VarPtr(ret(0)), VarPtr(.bfType), 2)
        Call MoveMemory(VarPtr(ret(2)), VarPtr(.bfSize), 4)
        Call MoveMemory(VarPtr(ret(6)), VarPtr(.bfReserved1), 2)
        Call MoveMemory(VarPtr(ret(8)), VarPtr(.bfReserved2), 2)
        Call MoveMemory(VarPtr(ret(10)), VarPtr(.bfOffBits), 4)
    End With

    With bih
        Call MoveMemory(VarPtr(ret(14)), VarPtr(.biSize), 4)
        Call MoveMemory(VarPtr(ret(18)), VarPtr(.biWidth), 4)
        Call MoveMemory(VarPtr(ret(22)), VarPtr(.biHeight), 4)
        Call MoveMemory(VarPtr(ret(26)), VarPtr(.biPlanes), 2)
        Call MoveMemory(VarPtr(ret(28)), VarPtr(.biBitCount), 2)
        Call MoveMemory(VarPtr(ret(30)), VarPtr(.biCompression), 4)
        Call MoveMemory(VarPtr(ret(34)), VarPtr(.biSizeImage), 4)
        Call MoveMemory(VarPtr(ret(38)), VarPtr(.biXPelsPerMeter), 4)
        Call MoveMemory(VarPtr(ret(42)), VarPtr(.biYPelsPerMeter), 4)
        Call MoveMemory(VarPtr(ret(46)), VarPtr(.biClrUsed), 4)
        Call MoveMemory(VarPtr(ret(50)), VarPtr(.biClrImportant), 4)
    End With

    Call MoveMemory(VarPtr(ret(54)), VarPtr(dataBlock(0)), UBound(dataBlock) + 1)
    
    Get24bppDIB = ret

End Function

'------------------------------------------------------------------------------
' (概要)
'  シンボルのイメージを返します。
'------------------------------------------------------------------------------
Public Function Get1bppImage(Optional ByVal moduleSize As Long = 5, _
                             Optional ByVal foreRGB As String = "#000000", _
                             Optional ByVal backRGB As String = "#FFFFFF") As IPicture

    If moduleSize < 1 Then Err.Raise 5
    
    Dim dib() As Byte
    dib = Get1bppDIB(moduleSize, foreRGB, backRGB)
    
    Dim ret As IPicture
    Set ret = ImageConverter.ConvertFrom(dib)

    Set Get1bppImage = ret
    
End Function

'------------------------------------------------------------------------------
' (概要)
'  シンボルのイメージを返します。
'------------------------------------------------------------------------------
Public Function Get24bppImage(Optional ByVal moduleSize As Long = 5, _
                              Optional ByVal foreRGB As String = "#000000", _
                              Optional ByVal backRGB As String = "#FFFFFF") As IPicture

    If moduleSize < 1 Then Err.Raise 5

    Dim dib() As Byte
    dib = Get24bppDIB(moduleSize, foreRGB, backRGB)
    
    Dim ret As IPicture
    Set ret = ImageConverter.ConvertFrom(dib)

    Set Get24bppImage = ret
    
End Function

'------------------------------------------------------------------------------
' (概要)
'  シンボルをビットマップ形式で保存します。
'------------------------------------------------------------------------------
Public Sub Save1bppDIB(ByVal filePath As String, _
                       Optional ByVal moduleSize As Long = 5, _
                       Optional ByVal foreRGB As String = "#000000", _
                       Optional ByVal backRGB As String = "#FFFFFF")
    
    If Len(filePath) = 0 Then Err.Raise 5
    If moduleSize < 1 Then Err.Raise 5

    If m_dataBitCounter = 0 Then Err.Raise 51

    Dim bitmapBytes() As Byte
    bitmapBytes = Get1bppDIB(moduleSize, foreRGB, backRGB)
    
    Dim fNum As Integer
    fNum = FreeFile()
    
    Open filePath For Binary As #fNum Len = &H7FFF
    Put #fNum, , bitmapBytes
    Close #fNum
    
End Sub

'------------------------------------------------------------------------------
' (概要)
'  シンボルをビットマップ形式で保存します。
'------------------------------------------------------------------------------
Public Sub Save24bppDIB(ByVal filePath As String, _
                        Optional ByVal moduleSize As Long = 5, _
                        Optional ByVal foreRGB As String = "#000000", _
                        Optional ByVal backRGB As String = "#FFFFFF")
    
    If Len(filePath) = 0 Then Err.Raise 5
    If moduleSize < 1 Then Err.Raise 5
    
    If m_dataBitCounter = 0 Then Err.Raise 51

    Dim bitmapBytes() As Byte
    bitmapBytes = Get24bppDIB(moduleSize, foreRGB, backRGB)
    
    Dim fNum As Integer
    fNum = FreeFile()
    
    Open filePath For Binary As #fNum Len = &H7FFF
    Put #fNum, , bitmapBytes
    Close #fNum
    
End Sub

