Attribute VB_Name = "ArrayUtil"
Option Private Module
Option Explicit

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
