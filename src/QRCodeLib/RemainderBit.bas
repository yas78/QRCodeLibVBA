Attribute VB_Name = "RemainderBit"
Option Private Module
Option Explicit

Public Sub Place(ByRef moduleMatrix() As Variant)
    Dim r As Long
    Dim c As Long
    For r = 0 To UBound(moduleMatrix)
        For c = 0 To UBound(moduleMatrix(r))
            If moduleMatrix(r)(c) = Values.BLANK Then
                moduleMatrix(r)(c) = -Values.WORD
            End If
        Next
    Next
End Sub
