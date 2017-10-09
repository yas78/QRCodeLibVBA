Attribute VB_Name = "QuietZone"
'----------------------------------------------------------------------------------------
' クワイエットゾーン
'----------------------------------------------------------------------------------------
Option Private Module
Option Explicit

Public Function Place(ByRef moduleMatrix() As Variant) As Variant()
    
    Const QUIET_ZONE_WIDTH As Integer = 4
    
    Dim ret() As Variant
    ReDim ret(UBound(moduleMatrix) + QUIET_ZONE_WIDTH * 2)
    
    Dim i As Long
    Dim cols() As Long
    
    For i = 0 To UBound(ret)
        ReDim cols(UBound(ret))
        ret(i) = cols
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
