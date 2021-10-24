Attribute VB_Name = "Separator"
Option Private Module
Option Explicit

Private Const VAL As Long = Values.SEPARATOR_PTN

Public Sub Place(ByRef moduleMatrix() As Variant)
    Dim offset As Long
    offset = UBound(moduleMatrix) - 7

    Dim i As Long
    For i = 0 To 7
         moduleMatrix(i)(7) = -VAL
         moduleMatrix(7)(i) = -VAL

         moduleMatrix(offset + i)(7) = -VAL
         moduleMatrix(offset + 0)(i) = -VAL

         moduleMatrix(i)(offset + 0) = -VAL
         moduleMatrix(7)(offset + i) = -VAL
     Next
End Sub
