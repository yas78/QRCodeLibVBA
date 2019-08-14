Attribute VB_Name = "ArrayUtil"
Option Private Module
Option Explicit


'------------------------------------------------------------------------------
' (概要)
'  二次元配列を左に90度回転した配列を返します。
'------------------------------------------------------------------------------
Public Function Rotate90(ByRef array2D() As Variant) As Variant()
    Dim ret() As Variant
    ReDim ret(UBound(array2D(0)))

    Dim i As Long
    Dim j As Long
    Dim rowArray() As Long

    For i = 0 To UBound(ret)
        ReDim rowArray(UBound(array2D))
        ret(i) = rowArray
    Next

    Dim k As Long
    k = UBound(ret)

    For i = 0 To UBound(ret)
        For j = 0 To UBound(ret(i))
            ret(i)(j) = array2D(j)(k - i)
        Next
    Next

    Rotate90 = ret
End Function

