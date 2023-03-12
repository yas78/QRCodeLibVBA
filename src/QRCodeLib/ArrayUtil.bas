Attribute VB_Name = "ArrayUtil"
Option Private Module
Option Explicit

Public Function Copy(ByRef dest() As Byte, ByVal destIdx As Long, _
                     ByRef src() As Byte, ByVal srcIdx As Long, _
                     ByVal sz As Long) As Long
    Dim idx As Long
    idx = destIdx

    Dim i As Long
    For i = 0 To sz - 1
        dest(idx) = src(srcIdx + i)
        idx = idx + 1
    Next

    Copy = idx
End Function

Public Function CopyAll(ByRef dest() As Byte, ByVal destIdx As Long, _
                        ByRef src() As Byte) As Long
    Dim sz As Long
    sz = UBound(src) + 1

    Dim ret As Long
    ret = Copy(dest, destIdx, src, 0, sz)

    CopyAll = ret
End Function

Public Function Rotate90(ByRef array2D() As Variant) As Variant()
    Dim ret() As Variant
    ReDim ret(UBound(array2D(0)))

    Dim rowArray() As Long

    Dim i1 As Long
    For i1 = 0 To UBound(ret)
        ReDim rowArray(UBound(array2D))
        ret(i1) = rowArray
    Next

    Dim k As Long
    k = UBound(ret)

    Dim i2 As Long
    Dim j  As Long
    For i2 = 0 To UBound(ret)
        For j = 0 To UBound(ret(i2))
            ret(i2)(j) = array2D(j)(k - i2)
        Next
    Next

    Rotate90 = ret
End Function

Public Function LongArray(ParamArray args() As Variant) As Long()
    Dim ret() As Long
    ReDim ret(UBound(args))

    Dim i As Long
    For i = 0 To UBound(args)
        ret(i) = args(i)
    Next

    LongArray = ret
End Function
