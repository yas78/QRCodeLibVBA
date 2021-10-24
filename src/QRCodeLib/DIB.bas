Attribute VB_Name = "DIB"
Option Private Module
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal pDest As LongPtr, ByVal pSrc As LongPtr, ByVal sz As Long)
#Else
    Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal pDest As Long, ByVal pSrc As Long, ByVal sz As Long)
#End If

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

    If Not monochrome Then
        biBitCount = 24
        bfOffBits = BF_SIZE + BI_SIZE
    Else
        ReDim palette(1)
        With palette(0)
            .rgbBlue = CByte((foreColorRgb And &HFF0000) \ 2 ^ 16)
            .rgbGreen = CByte((foreColorRgb And &HFF00&) \ 2 ^ 8)
            .rgbRed = CByte(foreColorRgb And &HFF&)
            .rgbReserved = 0
        End With

        With palette(1)
            .rgbBlue = CByte((backColorRgb And &HFF0000) \ 2 ^ 16)
            .rgbGreen = CByte((backColorRgb And &HFF00&) \ 2 ^ 8)
            .rgbRed = CByte(backColorRgb And &HFF&)
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

    With bfh
        Call MoveMemory(VarPtr(ret(0)), VarPtr(.bfType), 2)
        Call MoveMemory(VarPtr(ret(2)), VarPtr(.bfSize), 4)
        Call MoveMemory(VarPtr(ret(6)), VarPtr(.bfReserved1), 8)
    End With

    Call MoveMemory(VarPtr(ret(14)), VarPtr(bih.biSize), 40)

    Dim idx As Long
    idx = BF_SIZE + BI_SIZE
    Dim i As Long

    If monochrome Then
        For i = 0 To UBound(palette)
            Call MoveMemory(VarPtr(ret(idx)), VarPtr(palette(i)), 4)
            idx = idx + 4
        Next
    End If

    Call MoveMemory(VarPtr(ret(idx)), VarPtr(bitmapData(0)), UBound(bitmapData) + 1)

    GetDIB = ret
End Function
