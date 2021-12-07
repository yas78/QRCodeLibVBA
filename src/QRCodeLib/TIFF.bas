Attribute VB_Name = "TIFF"
Option Private Module
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal pDest As LongPtr, ByVal pSrc As LongPtr, ByVal sz As Long)
#Else
    Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal pDest As Long, ByVal pSrc As Long, ByVal sz As Long)
#End If

Public Enum TiffImageType
    Bilevel
    Grayscale
    PaletteColor
    FullColor
End Enum

Public Enum TiffFieldType
    [Byte] = 1
    Ascii = 2
    Short = 3
    [Long] = 4
    Rational = 5
End Enum

Public Enum TagID
    ImageWidth = &H100
    ImageLength = &H101
    BitsPerSample = &H102
    Compression = &H103
    PhotometricInterpretation = &H106
    StripOffsets = &H111
    SamplesPerPixel = &H115
    RowsPerStrip = &H116
    StripByteCounts = &H117
    XResolution = &H11A
    YResolution = &H11B
    ColorMap = &H140
End Enum

Private Enum TiffPhotometricInterpretation
    WhiteIsZero = 0
    BlackIsZero = 1
    RGB = 2
    PaletteColor = 3
End Enum

Private Type ImageFileHeader
    Data1 As Integer
    Data2 As Integer
    Data3 As Long
End Type

Private Type Rational
    Data1 As Long
    Data2 As Long
End Type

Private Type ColorPalette
    r(255) As Integer
    g(255) As Integer
    b(255) As Integer
End Type

Private m_imageType As TiffImageType

Private m_ifh As ImageFileHeader
Private m_ifd As ImageFileDirectory
Private m_data() As Byte

Private m_xResolution As Rational
Private m_yResolution As Rational
Private m_bitsPerSample(2) As Integer
Private m_palette As ColorPalette

Public Function GetTIFF(ByRef data() As Byte, _
                        ByVal pictWidth As Long, _
                        ByVal pictHeight As Long, _
                        ByVal imageType As TiffImageType, _
                        ByVal colors As Variant) As Byte()
    m_imageType = imageType

    m_ifh.Data1 = &H4949
    m_ifh.Data2 = 42
    m_ifh.Data3 = 8

    Set m_ifd = New ImageFileDirectory
    m_ifd.NextIFDOffset = 0

    m_data = data

    Call m_ifd.Add(TagID.ImageWidth, pictWidth)
    Call m_ifd.Add(TagID.ImageLength, pictHeight)

    Select Case imageType
        Case TiffImageType.Bilevel
            m_ifd.Add(TagID.BitsPerSample, 1).Count = 1
        Case TiffImageType.PaletteColor
            m_ifd.Add(TagID.BitsPerSample, 8).Count = 1
        Case TiffImageType.FullColor
            m_ifd.Add(TagID.BitsPerSample, 0).Count = 3
    End Select

    Call m_ifd.Add(TagID.Compression, 1)

    Select Case imageType
        Case TiffImageType.Bilevel
            Call m_ifd.Add(TagID.PhotometricInterpretation, TiffPhotometricInterpretation.BlackIsZero)
        Case TiffImageType.PaletteColor
            Call m_ifd.Add(TagID.PhotometricInterpretation, TiffPhotometricInterpretation.PaletteColor)
        Case TiffImageType.FullColor
            Call m_ifd.Add(TagID.PhotometricInterpretation, TiffPhotometricInterpretation.RGB)
    End Select

    Call m_ifd.Add(TagID.StripOffsets, 0)

    If imageType = TiffImageType.FullColor Then
        Call m_ifd.Add(TagID.SamplesPerPixel, 3)
    End If

    Call m_ifd.Add(TagID.RowsPerStrip, pictHeight)
    Call m_ifd.Add(TagID.StripByteCounts, UBound(data) + 1)
    Call m_ifd.Add(TagID.XResolution, 0)
    Call m_ifd.Add(TagID.YResolution, 0)

    If imageType = TiffImageType.PaletteColor Then
        m_ifd.Add(TagID.ColorMap, 0).Count = 3 * 2 ^ 8
    End If

    m_xResolution.Data1 = 96
    m_xResolution.Data2 = 1
    m_yResolution.Data1 = 96
    m_yResolution.Data2 = 1

    If imageType = TiffImageType.FullColor Then
        m_bitsPerSample(0) = 8
        m_bitsPerSample(1) = 8
        m_bitsPerSample(2) = 8
    End If

    If imageType = TiffImageType.PaletteColor Then
        Call MakeColorPalette(colors)
    End If

    Call UpdateOffsets

    Dim ret() As Byte
    Call ToBytes(ret)

    GetTIFF = ret
End Function

Private Sub MakeColorPalette(ByVal colors As Variant)
    Erase m_palette.r
    Erase m_palette.g
    Erase m_palette.b

    Dim i As Long
    For i = 0 To UBound(colors)
        m_palette.r(i) = ConvertTo16bits(colors(i) And &HFF&)
        m_palette.g(i) = ConvertTo16bits((colors(i) And &HFF00&) \ 2 ^ 8)
        m_palette.b(i) = ConvertTo16bits((colors(i) And &HFF0000) \ 2 ^ 16)
    Next
End Sub

Private Function ConvertTo16bits(ByVal arg As Byte) As Integer
    Dim temp As Long
    temp = arg
    temp = temp * ((2 ^ 16 - 1) \ &HFF)

    If (temp And &H8000&) > 0 Then
        temp = ((temp And &H7FFF&)) Or &H8000
    End If

    ConvertTo16bits = temp
End Function

Private Sub UpdateOffsets()
    Dim offset As Long
    offset = LenB(m_ifh) + m_ifd.Length

    Dim entries() As IFDEntry
    entries = m_ifd.GetEntries()

    Dim e As IFDEntry
    Dim i As Long
    For i = 0 To UBound(entries)
        Set e = entries(i)
        Select Case e.TagID
            Case TagID.BitsPerSample
                If m_imageType = TiffImageType.FullColor Then
                    e.Value = offset
                    offset = offset + e.GetDataSize()
                End If
            Case TagID.XResolution, TagID.YResolution, TagID.ColorMap
                e.Value = offset
                offset = offset + e.GetDataSize()
        End Select
    Next

    For i = 0 To UBound(entries)
        Set e = entries(i)
        If e.TagID = TagID.StripOffsets Then
            e.Value = offset
            Exit For
        End If
    Next
End Sub

Private Sub ToBytes(ByRef buffer() As Byte)
    ReDim buffer(CalcSize() - 1)

    Dim idx As Long
    idx = 0
    Dim sz As Long
    sz = 0

    sz = LenB(m_ifh)
    Call MoveMemory(VarPtr(buffer(idx)), VarPtr(m_ifh.Data1), sz)
    idx = idx + sz

    Dim ifdBytes() As Byte
    ifdBytes = m_ifd.GetBytes()
    sz = UBound(ifdBytes) + 1
    Call MoveMemory(VarPtr(buffer(idx)), VarPtr(ifdBytes(0)), sz)
    idx = idx + sz

    If m_imageType = TiffImageType.FullColor Then
        sz = 2 * (UBound(m_bitsPerSample) + 1)
        Call MoveMemory(VarPtr(buffer(idx)), VarPtr(m_bitsPerSample(0)), sz)
        idx = idx + sz
    End If

    sz = LenB(m_xResolution)
    Call MoveMemory(VarPtr(buffer(idx)), VarPtr(m_xResolution.Data1), sz)
    idx = idx + sz

    sz = LenB(m_yResolution)
    Call MoveMemory(VarPtr(buffer(idx)), VarPtr(m_yResolution.Data1), sz)
    idx = idx + sz

    If m_imageType = TiffImageType.PaletteColor Then
        sz = LenB(m_palette)
        Call MoveMemory(VarPtr(buffer(idx)), VarPtr(m_palette), sz)
        idx = idx + sz
    End If

    sz = UBound(m_data) + 1
    Call MoveMemory(VarPtr(buffer(idx)), VarPtr(m_data(0)), sz)
End Sub

Private Function CalcSize() As Long
    Dim ret As Long
    ret = LenB(m_ifh)
    ret = ret + m_ifd.Length

    If m_imageType = TiffImageType.FullColor Then
        ret = ret + 2 * (UBound(m_bitsPerSample) + 1)
    End If

    ret = ret + LenB(m_xResolution)
    ret = ret + LenB(m_yResolution)

    If m_imageType = TiffImageType.PaletteColor Then
        ret = ret + LenB(m_palette)
    End If

    ret = ret + (UBound(m_data) + 1)

    CalcSize = ret
End Function
