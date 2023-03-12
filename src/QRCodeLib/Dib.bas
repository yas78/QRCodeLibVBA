Attribute VB_Name = "Dib"
Option Private Module
Option Explicit

Private Const BF_SIZE As Long = 14
Private Const BI_SIZE As Long = 40

Private Type BitmapFileHeader
    bfType      As Integer
    bfSize      As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits   As Long
End Type

Private Type BitmapInfoHeader
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type

Private Type RgbQuad
    rgbBlue     As Byte
    rgbGreen    As Byte
    rgbRed      As Byte
    rgbReserved As Byte
End Type

Public Function GetDib(ByRef bitmapData() As Byte, _
                       ByVal pictWidth As Long, _
                       ByVal pictHeight As Long, _
                       ByVal foreColorRgb As Long, _
                       ByVal backColorRgb As Long, _
                       ByVal monochrome As Boolean) As Byte()
    Dim bfOffBits  As Long
    Dim biBitCount As Integer

    Dim bfh As BitmapFileHeader
    Dim bih As BitmapInfoHeader
    Dim palette() As RgbQuad

    Dim bytes() As Byte

    If Not monochrome Then
        biBitCount = 24
        bfOffBits = BF_SIZE + BI_SIZE
    Else
        ReDim palette(1)

        bytes = BitConverter.GetBytes(foreColorRgb)
        With palette(0)
            .rgbBlue = bytes(2)
            .rgbGreen = bytes(1)
            .rgbRed = bytes(0)
            .rgbReserved = 0
        End With

        bytes = BitConverter.GetBytes(backColorRgb)
        With palette(1)
            .rgbBlue = bytes(2)
            .rgbGreen = bytes(1)
            .rgbRed = bytes(0)
            .rgbReserved = 0
        End With

        biBitCount = 1
        bfOffBits = BF_SIZE + BI_SIZE + (4 * (UBound(palette) + 1))
    End If

    With bfh
        .bfType = &H4D42
        .bfSize = bfOffBits + (UBound(bitmapData) + 1)
        .bfReserved1 = 0
        .bfReserved2 = 0
        .bfOffBits = bfOffBits
    End With

    With bih
        .biSize = BI_SIZE
        .biWidth = pictWidth
        .biHeight = pictHeight
        .biPlanes = 1
        .biBitCount = biBitCount
        .biCompression = 0
        .biSizeImage = 0
        .biXPelsPerMeter = 0
        .biYPelsPerMeter = 0
        .biClrUsed = 0
        .biClrImportant = 0
    End With

    Dim bs As New ByteSequence

    With bfh
        Call bs.Append(.bfType)
        Call bs.Append(.bfSize)
        Call bs.Append(.bfReserved1)
        Call bs.Append(.bfReserved2)
        Call bs.Append(.bfOffBits)
    End With

    With bih
        Call bs.Append(.biSize)
        Call bs.Append(.biWidth)
        Call bs.Append(.biHeight)
        Call bs.Append(.biPlanes)
        Call bs.Append(.biBitCount)
        Call bs.Append(.biCompression)
        Call bs.Append(.biSizeImage)
        Call bs.Append(.biXPelsPerMeter)
        Call bs.Append(.biYPelsPerMeter)
        Call bs.Append(.biClrUsed)
        Call bs.Append(.biClrImportant)
    End With

    Dim i As Long

    If monochrome Then
        For i = 0 To UBound(palette)
            Call bs.Append(palette(i).rgbBlue)
            Call bs.Append(palette(i).rgbGreen)
            Call bs.Append(palette(i).rgbRed)
            Call bs.Append(palette(i).rgbReserved)
        Next
    End If

    Call bs.Append(bitmapData)

    GetDib = bs.Flush()
End Function
