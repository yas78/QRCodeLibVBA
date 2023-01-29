Attribute VB_Name = "Zlib"
Option Private Module
Option Explicit

Private Enum CompressionLevel
    Fastest = 0
    Fast = 1
    Default = 2
    Slowest = 3
End Enum

Public Sub Compress(ByRef data() As Byte, ByVal btype As DeflateBType, ByRef buffer() As Byte)
    Dim cmf As Byte
    cmf = &H78

    Dim fdict As Byte
    fdict = &H0

    Dim flevel As Byte
    flevel = CompressionLevel.Default * 2 ^ 6

    Dim flg As Byte
    flg = flevel + fdict

    Dim fcheck As Byte
    fcheck = 31 - ((cmf * 2 ^ 8 + flg) Mod 31)

    flg = flg + fcheck

    Dim compressedData() As Byte
    Call Deflate.Compress(data, btype, compressedData)

    Dim adler As Long
    adler = Adler32.Checksum(data)

    Dim sz As Long
    sz = UBound(compressedData) + 1

    ReDim buffer(1 + 1 + sz + 4 - 1)

    Dim idx As Long
    idx = 0

    buffer(idx) = cmf
    idx = idx + 1

    buffer(idx) = flg
    idx = idx + 1

    idx = ArrayUtil.CopyAll(buffer, idx, compressedData)

    Dim bytes() As Byte
    bytes = BitConverter.GetBytes(adler, True)
    Call ArrayUtil.CopyAll(buffer, idx, bytes)
End Sub
