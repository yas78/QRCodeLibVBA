Attribute VB_Name = "QuietZone"
'------------------------------------------------------------------------------
' クワイエットゾーン
'------------------------------------------------------------------------------
Option Private Module
Option Explicit


Public Const QUIET_ZONE_WIDTH As Long = 4

Public Function Place(ByRef moduleMatrix() As Variant) As Variant()
    Dim sz As Long
    sz = UBound(moduleMatrix) + QUIET_ZONE_WIDTH * 2

    Dim ret() As Variant
    ReDim ret(sz)

    Dim i As Long
    Dim rowArray() As Long

    For i = 0 To sz
        ReDim rowArray(sz)
        ret(i) = rowArray
    Next

    Dim r As Long
    Dim c As Long

    For r = 0 To UBound(moduleMatrix)
        For c = 0 To UBound(moduleMatrix(r))
            ret(r + QUIET_ZONE_WIDTH)(c + QUIET_ZONE_WIDTH) = moduleMatrix(r)(c)
        Next
    Next

    Place = ret
End Function
