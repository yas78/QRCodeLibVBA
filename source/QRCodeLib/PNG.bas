Attribute VB_Name = "PNG"
Option Private Module
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal pDest As LongPtr, ByVal pSrc As LongPtr, ByVal sz As Long)
#Else
    Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal pDest As Long, ByVal pSrc As Long, ByVal sz As Long)
#End If

Public Type PngSignature
    psData(7) As Byte
End Type

Public Type PngChunk
    pLength As Long
    pType   As Long
    pData() As Byte
    pCRC    As Long
End Type

Private Enum PngColorType
    pGrayscale = 0
    pTrueColor = 2
    pIndexColor = 3
    pGrayscaleAlpha = 4
    pTrueColorAlpha = 6
End Enum

Private Enum PngCompressionMethod
    Deflate = 0
End Enum

Private Enum PngFilterType
    pNone = 0
    pSub = 1
    pUp = 2
    pAverage = 3
    pPaeth = 4
End Enum

Private Enum PngInterlaceMethod
    pNo = 0
    pAdam7 = 1
End Enum

Public Function BuildMonochromeBin(ByRef data() As Byte, _
                                   ByVal pictWidth As Long, _
                                   ByVal pictHeight As Long, _
                                   ByVal foreColorRgb As Long, _
                                   ByVal backColorRGB As Long) As Byte()
    Dim psgn As PngSignature
    Call MakePngSignature(psgn)

    Dim ihdr As PngChunk
    Call MakeIHDR( _
        pictWidth, _
        pictHeight, _
        1, _
        PngColorType.pIndexColor, _
        PngCompressionMethod.Deflate, _
        PngFilterType.pNone, _
        PngInterlaceMethod.pNo, _
        ihdr _
    )

    Dim iplt As PngChunk
    Call MakeIPLT(iplt, foreColorRgb, backColorRGB)

    Dim idat As PngChunk
    Call MakeIDAT(data, DeflateBType.NoCompression, idat)

    Dim iend As PngChunk
    Call MakeIEND(iend)

    Dim ret() As Byte
    Call ToBytes(psgn, ihdr, iplt, idat, iend, ret)

    BuildMonochromeBin = ret
End Function

Public Function BuildTrueColorBin(ByRef data() As Byte, _
                                  ByVal pictWidth As Long, _
                                  ByVal pictHeight As Long) As Byte()
    Dim psgn As PngSignature
    Call MakePngSignature(psgn)

    Dim ihdr As PngChunk
    Call MakeIHDR( _
        pictWidth, _
        pictHeight, _
        8, _
        PngColorType.pTrueColor, _
        PngCompressionMethod.Deflate, _
        PngFilterType.pNone, _
        PngInterlaceMethod.pNo, _
        ihdr _
    )

    Dim iplt As PngChunk

    Dim idat As PngChunk
    Call MakeIDAT(data, DeflateBType.NoCompression, idat)

    Dim iend As PngChunk
    Call MakeIEND(iend)

    Dim ret() As Byte
    Call ToBytes(psgn, ihdr, iplt, idat, iend, ret)

    BuildTrueColorBin = ret
End Function

Private Sub MakePngSignature(ByRef psgn As PngSignature)
    With psgn
        .psData(0) = &H89
        .psData(1) = Asc("P")
        .psData(2) = Asc("N")
        .psData(3) = Asc("G")
        .psData(4) = Asc(vbCr)
        .psData(5) = Asc(vbLf)
        .psData(6) = &H1A
        .psData(7) = Asc(vbLf)
    End With
End Sub

Private Sub MakeIHDR(ByVal pictWidth As Long, _
                     ByVal pictHeight As Long, _
                     ByVal bitDepth As Long, _
                     ByVal tColor As PngColorType, _
                     ByVal compression As Long, _
                     ByVal tFilter As Long, _
                     ByVal interlace As Long, _
                     ByRef ihdr As PngChunk)
    Const STR_IHDR As Long = &H49484452

    Dim lbe As Long
    Dim crc As Long

    With ihdr
        .pLength = 13
        .pType = STR_IHDR

        ReDim .pData(.pLength - 1)
        lbe = BitConverter.ToBigEndian(pictWidth)
        Call MoveMemory(VarPtr(.pData(0)), VarPtr(lbe), 4)
        lbe = BitConverter.ToBigEndian(pictHeight)
        Call MoveMemory(VarPtr(.pData(4)), VarPtr(lbe), 4)

        .pData(8) = bitDepth
        .pData(9) = tColor
        .pData(10) = compression
        .pData(11) = tFilter
        .pData(12) = interlace

        crc = CRC32.Checksum(BitConverter.GetBytes(STR_IHDR, True))
        .pCRC = CRC32.Update(crc, .pData)
    End With
End Sub

Private Sub MakeIPLT(ByRef iplt As PngChunk, ParamArray rgbArray() As Variant)
    Const STR_PLTE As Long = &H504C5445

    Dim idx As Long
    idx = 0

    Dim v   As Variant
    Dim crc As Long

    With iplt
        .pLength = (UBound(rgbArray) + 1) * 3
        .pType = STR_PLTE

        ReDim .pData(.pLength - 1)
        For Each v In rgbArray
            .pData(idx + 0) = CByte(v And &HFF&)
            .pData(idx + 1) = CByte((v And &HFF00&) \ 2 ^ 8)
            .pData(idx + 2) = CByte((v And &HFF0000) \ 2 ^ 16)
            idx = idx + 3
        Next

        crc = CRC32.Checksum(BitConverter.GetBytes(.pType, True))
        .pCRC = CRC32.Update(crc, .pData)
    End With
End Sub

Private Sub MakeIDAT(ByRef data() As Byte, ByVal btype As DeflateBType, ByRef idat As PngChunk)
    Const STR_IDAT As Long = &H49444154

    Dim crc As Long

    With idat
        Call ZLIB.Compress(data, btype, .pData)
        .pLength = UBound(.pData) + 1
        .pType = STR_IDAT
        crc = CRC32.Checksum(BitConverter.GetBytes(STR_IDAT, True))
        .pCRC = CRC32.Update(crc, .pData)
    End With
End Sub

Private Sub MakeIEND(ByRef iend As PngChunk)
    Const STR_IEND As Long = &H49454E44

    With iend
        .pLength = 0
        .pType = STR_IEND
        .pCRC = CRC32.Checksum(BitConverter.GetBytes(STR_IEND, True))
    End With
End Sub

Private Sub ToBytes(ByRef psgn As PngSignature, _
                    ByRef ihdr As PngChunk, _
                    ByRef iplt As PngChunk, _
                    ByRef idat As PngChunk, _
                    ByRef iend As PngChunk, _
                    ByRef buffer() As Byte)
    Dim pfhSize As Long
    pfhSize = 8

    Dim ihdrSize As Long
    ihdrSize = 12 + ihdr.pLength

    Dim ipltSize As Long

    If iplt.pLength > 0 Then
        ipltSize = 12 + iplt.pLength
    Else
        ipltSize = 0
    End If

    Dim idatSize As Long
    idatSize = 12 + idat.pLength

    Dim iendSize As Long
    iendSize = 12 + iend.pLength

    ReDim buffer(pfhSize + ihdrSize + ipltSize + idatSize + iendSize - 1)

    Dim idx As Long
    idx = 0

    With psgn
        Call MoveMemory(VarPtr(buffer(idx)), VarPtr(.psData(0)), 8)
        idx = idx + 8
    End With

    Dim lbe As Long

    With ihdr
        lbe = BitConverter.ToBigEndian(.pLength)
        Call MoveMemory(VarPtr(buffer(idx)), VarPtr(lbe), 4)
        idx = idx + 4
        lbe = BitConverter.ToBigEndian(.pType)
        Call MoveMemory(VarPtr(buffer(idx)), VarPtr(lbe), 4)
        idx = idx + 4
        Call MoveMemory(VarPtr(buffer(idx)), VarPtr(.pData(0)), .pLength)
        idx = idx + .pLength
        lbe = BitConverter.ToBigEndian(.pCRC)
        Call MoveMemory(VarPtr(buffer(idx)), VarPtr(lbe), 4)
        idx = idx + 4
    End With

    If iplt.pLength > 0 Then
        With iplt
            lbe = BitConverter.ToBigEndian(.pLength)
            Call MoveMemory(VarPtr(buffer(idx)), VarPtr(lbe), 4)
            idx = idx + 4
            lbe = BitConverter.ToBigEndian(.pType)
            Call MoveMemory(VarPtr(buffer(idx)), VarPtr(lbe), 4)
            idx = idx + 4
            Call MoveMemory(VarPtr(buffer(idx)), VarPtr(.pData(0)), .pLength)
            idx = idx + .pLength
            lbe = BitConverter.ToBigEndian(.pCRC)
            Call MoveMemory(VarPtr(buffer(idx)), VarPtr(lbe), 4)
            idx = idx + 4
        End With
    End If

    With idat
        lbe = BitConverter.ToBigEndian(.pLength)
        Call MoveMemory(VarPtr(buffer(idx)), VarPtr(lbe), 4)
        idx = idx + 4
        lbe = BitConverter.ToBigEndian(.pType)
        Call MoveMemory(VarPtr(buffer(idx)), VarPtr(lbe), 4)
        idx = idx + 4
        Call MoveMemory(VarPtr(buffer(idx)), VarPtr(.pData(0)), .pLength)
        idx = idx + .pLength
        lbe = BitConverter.ToBigEndian(.pCRC)
        Call MoveMemory(VarPtr(buffer(idx)), VarPtr(lbe), 4)
        idx = idx + 4
    End With

    With iend
        lbe = BitConverter.ToBigEndian(.pLength)
        Call MoveMemory(VarPtr(buffer(idx)), VarPtr(lbe), 4)
        idx = idx + 4
        lbe = BitConverter.ToBigEndian(.pType)
        Call MoveMemory(VarPtr(buffer(idx)), VarPtr(lbe), 4)
        idx = idx + 4
        lbe = BitConverter.ToBigEndian(.pCRC)
        Call MoveMemory(VarPtr(buffer(idx)), VarPtr(lbe), 4)
    End With
End Sub
