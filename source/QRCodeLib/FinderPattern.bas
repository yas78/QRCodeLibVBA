Attribute VB_Name = "FinderPattern"
'---------------------------------------------------------------------------
' 位置検出パターン
'---------------------------------------------------------------------------
Option Private Module
Option Explicit


Private m_finderPattern() As Variant
Private m_initialized As Boolean

'--------------------------------------------------------------------------------
' (概要)
'  位置検出パターンを配置します。
'--------------------------------------------------------------------------------
Public Sub Place(ByRef moduleMatrix() As Variant)
    Call Init

    Dim offset As Long
    offset = (UBound(moduleMatrix) + 1) - (UBound(m_finderPattern) + 1)

    Dim i As Long
    Dim j As Long
    Dim v As Long

    For i = 0 To UBound(m_finderPattern)
        For j = 0 To UBound(m_finderPattern(i))
            v = m_finderPattern(i)(j)

            moduleMatrix(i)(j) = v
            moduleMatrix(i)(j + offset) = v
            moduleMatrix(i + offset)(j) = v
        Next
    Next
End Sub

'------------------------------------------------------------------------------
' (概要)
'  オブジェクトを初期化します。
'------------------------------------------------------------------------------
Private Sub Init()
    If m_initialized Then Exit Sub

    m_initialized = True

   ' 位置検出パターン
    m_finderPattern = Array( _
        Array(2, 2, 2, 2, 2, 2, 2), _
        Array(2, -2, -2, -2, -2, -2, 2), _
        Array(2, -2, 2, 2, 2, -2, 2), _
        Array(2, -2, 2, 2, 2, -2, 2), _
        Array(2, -2, 2, 2, 2, -2, 2), _
        Array(2, -2, -2, -2, -2, -2, 2), _
        Array(2, 2, 2, 2, 2, 2, 2) _
    )
End Sub
