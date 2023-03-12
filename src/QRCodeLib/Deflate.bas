Attribute VB_Name = "Deflate"
Option Private Module
Option Explicit

Public Enum DeflateBType
    NoCompression = 0
    CompressedWithFixedHuffmanCodes = 1
    CompressedWithDynamicHuffmanCodes = 2
    Reserved = 3
End Enum

Public Sub Compress(ByRef data() As Byte, ByVal btype As DeflateBType, ByRef buffer() As Byte)
    If btype <> DeflateBType.NoCompression Then Call Err.Raise(5)

    Dim bytesLen As Long
    bytesLen = UBound(data) + 1

    Dim quotient As Long
    quotient = bytesLen \ &HFFFF&

    Dim remainder As Long
    remainder = bytesLen Mod &HFFFF&

    Dim bufferSize As Long
    bufferSize = quotient * (1 + 4 + &HFFFF&)

    If remainder > 0 Then
        bufferSize = bufferSize + (1 + 4 + remainder)
    End If

    ReDim buffer(bufferSize - 1)

    Dim bfinal As Long
    Dim dLen   As Long
    Dim dNLen  As Long

    Dim idx As Long
    idx = 0

    Dim bytes() As Byte

    Dim i As Long
    For i = 0 To quotient - 1
        bfinal = 0
        buffer(idx) = bfinal Or (btype * 2 ^ 1)
        idx = idx + 1

        dLen = &HFFFF&
        bytes = BitConverter.GetBytes(dLen)
        idx = ArrayUtil.Copy(buffer, idx, bytes, 0, 2)

        dNLen = dLen Xor &HFFFF&
        bytes = BitConverter.GetBytes(dNLen)
        idx = ArrayUtil.Copy(buffer, idx, bytes, 0, 2)

        idx = ArrayUtil.Copy(buffer, idx, data, &HFFFF& * i, &HFFFF&)
    Next

    If remainder > 0 Then
        bfinal = 1
        buffer(idx) = bfinal Or (btype * 2 ^ 1)
        idx = idx + 1

        dLen = remainder
        bytes = BitConverter.GetBytes(dLen)
        idx = ArrayUtil.Copy(buffer, idx, bytes, 0, 2)

        dNLen = dLen Xor &HFFFF&
        bytes = BitConverter.GetBytes(dNLen)
        idx = ArrayUtil.Copy(buffer, idx, bytes, 0, 2)

        idx = ArrayUtil.Copy(buffer, idx, data, &HFFFF& * quotient, remainder)
    End If
End Sub
