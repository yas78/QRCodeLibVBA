Attribute VB_Name = "TimingPattern"
'------------------------------------------------------------------------------
' タイミングパターン
'------------------------------------------------------------------------------
Option Private Module
Option Explicit

'------------------------------------------------------------------------------
' (概要)
'  タイミングパターンを配置します。
'------------------------------------------------------------------------------
Public Sub Place(ByRef moduleMatrix() As Variant)

    Dim i As Long
    Dim v As Long

    For i = 8 To UBound(moduleMatrix) - 8
        v = IIf(i Mod 2 = 0, 2, -2)

        moduleMatrix(6)(i) = v
        moduleMatrix(i)(6) = v
    Next

End Sub
