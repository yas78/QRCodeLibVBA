Attribute VB_Name = "AlignmentPattern"
Option Private Module
Option Explicit

Private m_centerPosArrays(40) As Variant

Private m_initialized As Boolean

'------------------------------------------------------------------------------
' (概要)
'  位置合わせパターンを配置します。
'------------------------------------------------------------------------------
Public Sub Place(ByRef moduleMatrix() As Variant, ByVal ver As Long)

#If [DEBUG] Then
    Debug.Assert ver >= 2 And ver <= Constants.MAX_VERSION
#End If

    Call Initialize

    Dim centerArray As Variant
    centerArray = m_centerPosArrays(ver)

    Dim maxIndex As Long
    maxIndex = UBound(centerArray)
    
    Dim i As Long
    Dim j As Long
    
    Dim r As Long
    Dim c As Long
    
    For i = 0 To maxIndex
        r = centerArray(i)

        For j = 0 To maxIndex
            c = centerArray(j)

            ' 位置検出パターンと重なる場合
            If i = 0 And j = 0 Or _
               i = 0 And j = maxIndex Or _
               i = maxIndex And j = 0 Then
                
                GoTo Continue_j
                
            End If
            
            moduleMatrix(r - 2)(c - 2) = 2
            moduleMatrix(r - 2)(c - 1) = 2
            moduleMatrix(r - 2)(c + 0) = 2
            moduleMatrix(r - 2)(c + 1) = 2
            moduleMatrix(r - 2)(c + 2) = 2
            
            moduleMatrix(r - 1)(c - 2) = 2
            moduleMatrix(r - 1)(c - 1) = -2
            moduleMatrix(r - 1)(c + 0) = -2
            moduleMatrix(r - 1)(c + 1) = -2
            moduleMatrix(r - 1)(c + 2) = 2
            
            moduleMatrix(r + 0)(c - 2) = 2
            moduleMatrix(r + 0)(c - 1) = -2
            moduleMatrix(r + 0)(c + 0) = 2
            moduleMatrix(r + 0)(c + 1) = -2
            moduleMatrix(r + 0)(c + 2) = 2
            
            moduleMatrix(r + 1)(c - 2) = 2
            moduleMatrix(r + 1)(c - 1) = -2
            moduleMatrix(r + 1)(c + 0) = -2
            moduleMatrix(r + 1)(c + 1) = -2
            moduleMatrix(r + 1)(c + 2) = 2
            
            moduleMatrix(r + 2)(c - 2) = 2
            moduleMatrix(r + 2)(c - 1) = 2
            moduleMatrix(r + 2)(c + 0) = 2
            moduleMatrix(r + 2)(c + 1) = 2
            moduleMatrix(r + 2)(c + 2) = 2
            
Continue_j:
        Next
    Next
    
End Sub

'------------------------------------------------------------------------------
' (概要)
'  オブジェクトを初期化します。
'------------------------------------------------------------------------------
Private Sub Initialize()

    If m_initialized Then Exit Sub

    m_initialized = True

    ' 位置合せパターンの中心座標
    m_centerPosArrays(2) = Array(6, 18)
    m_centerPosArrays(3) = Array(6, 22)
    m_centerPosArrays(4) = Array(6, 26)
    m_centerPosArrays(5) = Array(6, 30)
    m_centerPosArrays(6) = Array(6, 34)
    m_centerPosArrays(7) = Array(6, 22, 38)
    m_centerPosArrays(8) = Array(6, 24, 42)
    m_centerPosArrays(9) = Array(6, 26, 46)
    m_centerPosArrays(10) = Array(6, 28, 50)
    m_centerPosArrays(11) = Array(6, 30, 54)
    m_centerPosArrays(12) = Array(6, 32, 58)
    m_centerPosArrays(13) = Array(6, 34, 62)
    m_centerPosArrays(14) = Array(6, 26, 46, 66)
    m_centerPosArrays(15) = Array(6, 26, 48, 70)
    m_centerPosArrays(16) = Array(6, 26, 50, 74)
    m_centerPosArrays(17) = Array(6, 30, 54, 78)
    m_centerPosArrays(18) = Array(6, 30, 56, 82)
    m_centerPosArrays(19) = Array(6, 30, 58, 86)
    m_centerPosArrays(20) = Array(6, 34, 62, 90)
    m_centerPosArrays(21) = Array(6, 28, 50, 72, 94)
    m_centerPosArrays(22) = Array(6, 26, 50, 74, 98)
    m_centerPosArrays(23) = Array(6, 30, 54, 78, 102)
    m_centerPosArrays(24) = Array(6, 28, 54, 80, 106)
    m_centerPosArrays(25) = Array(6, 32, 58, 84, 110)
    m_centerPosArrays(26) = Array(6, 30, 58, 86, 114)
    m_centerPosArrays(27) = Array(6, 34, 62, 90, 118)
    m_centerPosArrays(28) = Array(6, 26, 50, 74, 98, 122)
    m_centerPosArrays(29) = Array(6, 30, 54, 78, 102, 126)
    m_centerPosArrays(30) = Array(6, 26, 52, 78, 104, 130)
    m_centerPosArrays(31) = Array(6, 30, 56, 82, 108, 134)
    m_centerPosArrays(32) = Array(6, 34, 60, 86, 112, 138)
    m_centerPosArrays(33) = Array(6, 30, 58, 86, 114, 142)
    m_centerPosArrays(34) = Array(6, 34, 62, 90, 118, 146)
    m_centerPosArrays(35) = Array(6, 30, 54, 78, 102, 126, 150)
    m_centerPosArrays(36) = Array(6, 24, 50, 76, 102, 128, 154)
    m_centerPosArrays(37) = Array(6, 28, 54, 80, 106, 132, 158)
    m_centerPosArrays(38) = Array(6, 32, 58, 84, 110, 136, 162)
    m_centerPosArrays(39) = Array(6, 26, 54, 82, 110, 138, 166)
    m_centerPosArrays(40) = Array(6, 30, 58, 86, 114, 142, 170)
    
End Sub
