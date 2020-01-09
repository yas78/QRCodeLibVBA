Attribute VB_Name = "Separator"
Option Private Module
Option Explicit

Public Sub Place(ByRef moduleMatrix() As Variant)
    Dim offset As Long
    offset = UBound(moduleMatrix) - 7

    Dim i As Long

    For i = 0 To 7
         moduleMatrix(i)(7) = -2
         moduleMatrix(7)(i) = -2

         moduleMatrix(offset + i)(7) = -2
         moduleMatrix(offset + 0)(i) = -2

         moduleMatrix(i)(offset + 0) = -2
         moduleMatrix(7)(offset + i) = -2
     Next
End Sub
