Attribute VB_Name = "RemainderBit"
'------------------------------------------------------------------------------
' 残余ビット
'------------------------------------------------------------------------------
Option Private Module
Option Explicit

Public Sub Place(ByRef moduleMatrix() As Variant)

    Dim r As Long
    Dim c As Long

    For r = 0 To UBound(moduleMatrix)
        For c = 0 To UBound(moduleMatrix(r))
            If moduleMatrix(r)(c) = 0 Then
                moduleMatrix(r)(c) = -1
            End If
        Next
    Next

End Sub
