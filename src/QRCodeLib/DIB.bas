Attribute VB_Name = "DIB"
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

Public Function GetDIB(ByRef bitmapData() As Byte, _
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

    Dim ret() As Byte
    ReDim ret(bfOffBits + UBound(bitmapData))

    Dim idx As Long
    idx = 0

    With bfh
        bytes = BitConverter.GetBytes(.bfType)
        idx = ArrayUtil.CopyAll(ret, idx, bytes)
        bytes = BitConverter.GetBytes(.bfSize)
        idx = ArrayUtil.CopyAll(ret, idx, bytes)
        bytes = BitConverter.GetBytes(.bfReserved1)
        idx = ArrayUtil.CopyAll(ret, idx, bytes)
        bytes = BitConverter.GetBytes(.bfReserved2)
        idx = ArrayUtil.CopyAll(ret, idx, bytes)
        bytes = BitConverter.GetBytes(.bfOffBits)
        idx = ArrayUtil.CopyAll(ret, idx, bytes)
    End With

    With bih
        bytes = BitConverter.GetBytes(.biSize)
        idx = ArrayUtil.CopyAll(ret, idx, bytes)
        bytes = BitConverter.GetBytes(.biWidth)
        idx = ArrayUtil.CopyAll(ret, idx, bytes)
        bytes = BitConverter.GetBytes(.biHeight)
        idx = ArrayUtil.CopyAll(ret, idx, bytes)
        bytes = BitConverter.GetBytes(.biPlanes)
        idx = ArrayUtil.CopyAll(ret, idx, bytes)
        bytes = BitConverter.GetBytes(.biBitCount)
        idx = ArrayUtil.CopyAll(ret, idx, bytes)
        bytes = BitConverter.GetBytes(.biCompression)
        idx = ArrayUtil.CopyAll(ret, idx, bytes)
        bytes = BitConverter.GetBytes(.biSizeImage)
        idx = ArrayUtil.CopyAll(ret, idx, bytes)
        bytes = BitConverter.GetBytes(.biXPelsPerMeter)
        idx = ArrayUtil.CopyAll(ret, idx, bytes)
        bytes = BitConverter.GetBytes(.biYPelsPerMeter)
        idx = ArrayUtil.CopyAll(ret, idx, bytes)
        bytes = BitConverter.GetBytes(.biClrUsed)
        idx = ArrayUtil.CopyAll(ret, idx, bytes)
        bytes = BitConverter.GetBytes(.biClrImportant)
        idx = ArrayUtil.CopyAll(ret, idx, bytes)
    End With

    Dim i As Long

    If monochrome Then
        For i = 0 To UBound(palette)
            ret(idx + 0) = palette(i).rgbBlue
            ret(idx + 1) = palette(i).rgbGreen
            ret(idx + 2) = palette(i).rgbRed
            ret(idx + 3) = palette(i).rgbReserved
            idx = idx + 4
        Next
    End If

    Call ArrayUtil.CopyAll(ret, idx, bitmapData)

    GetDIB = ret
End Function
