VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AlignmentPattern"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const VAL As Long = Values.ALIGNMENT_PTN

Private m_items(40) As Variant

Private Sub Class_Initialize()
    m_items(2) = LongArray(6, 18)
    m_items(3) = LongArray(6, 22)
    m_items(4) = LongArray(6, 26)
    m_items(5) = LongArray(6, 30)
    m_items(6) = LongArray(6, 34)
    m_items(7) = LongArray(6, 22, 38)
    m_items(8) = LongArray(6, 24, 42)
    m_items(9) = LongArray(6, 26, 46)
    m_items(10) = LongArray(6, 28, 50)
    m_items(11) = LongArray(6, 30, 54)
    m_items(12) = LongArray(6, 32, 58)
    m_items(13) = LongArray(6, 34, 62)
    m_items(14) = LongArray(6, 26, 46, 66)
    m_items(15) = LongArray(6, 26, 48, 70)
    m_items(16) = LongArray(6, 26, 50, 74)
    m_items(17) = LongArray(6, 30, 54, 78)
    m_items(18) = LongArray(6, 30, 56, 82)
    m_items(19) = LongArray(6, 30, 58, 86)
    m_items(20) = LongArray(6, 34, 62, 90)
    m_items(21) = LongArray(6, 28, 50, 72, 94)
    m_items(22) = LongArray(6, 26, 50, 74, 98)
    m_items(23) = LongArray(6, 30, 54, 78, 102)
    m_items(24) = LongArray(6, 28, 54, 80, 106)
    m_items(25) = LongArray(6, 32, 58, 84, 110)
    m_items(26) = LongArray(6, 30, 58, 86, 114)
    m_items(27) = LongArray(6, 34, 62, 90, 118)
    m_items(28) = LongArray(6, 26, 50, 74, 98, 122)
    m_items(29) = LongArray(6, 30, 54, 78, 102, 126)
    m_items(30) = LongArray(6, 26, 52, 78, 104, 130)
    m_items(31) = LongArray(6, 30, 56, 82, 108, 134)
    m_items(32) = LongArray(6, 34, 60, 86, 112, 138)
    m_items(33) = LongArray(6, 30, 58, 86, 114, 142)
    m_items(34) = LongArray(6, 34, 62, 90, 118, 146)
    m_items(35) = LongArray(6, 30, 54, 78, 102, 126, 150)
    m_items(36) = LongArray(6, 24, 50, 76, 102, 128, 154)
    m_items(37) = LongArray(6, 28, 54, 80, 106, 132, 158)
    m_items(38) = LongArray(6, 32, 58, 84, 110, 136, 162)
    m_items(39) = LongArray(6, 26, 54, 82, 110, 138, 166)
    m_items(40) = LongArray(6, 30, 58, 86, 114, 142, 170)
End Sub

Public Sub Place(ByVal ver As Long, ByRef moduleMatrix() As Variant)
    Dim centers() As Long
    centers = m_items(ver)

    Dim maxIndex As Long
    maxIndex = UBound(centers)

    Dim i As Long
    Dim j As Long
    Dim r As Long
    Dim c As Long
    For i = 0 To maxIndex
        r = centers(i)

        For j = 0 To maxIndex
            c = centers(j)

            If i = 0 And j = 0 Or _
               i = 0 And j = maxIndex Or _
               i = maxIndex And j = 0 Then

                GoTo Continue
            End If

            moduleMatrix(r - 2)(c - 2) = VAL
            moduleMatrix(r - 2)(c - 1) = VAL
            moduleMatrix(r - 2)(c + 0) = VAL
            moduleMatrix(r - 2)(c + 1) = VAL
            moduleMatrix(r - 2)(c + 2) = VAL

            moduleMatrix(r - 1)(c - 2) = VAL
            moduleMatrix(r - 1)(c - 1) = -VAL
            moduleMatrix(r - 1)(c + 0) = -VAL
            moduleMatrix(r - 1)(c + 1) = -VAL
            moduleMatrix(r - 1)(c + 2) = VAL

            moduleMatrix(r + 0)(c - 2) = VAL
            moduleMatrix(r + 0)(c - 1) = -VAL
            moduleMatrix(r + 0)(c + 0) = VAL
            moduleMatrix(r + 0)(c + 1) = -VAL
            moduleMatrix(r + 0)(c + 2) = VAL

            moduleMatrix(r + 1)(c - 2) = VAL
            moduleMatrix(r + 1)(c - 1) = -VAL
            moduleMatrix(r + 1)(c + 0) = -VAL
            moduleMatrix(r + 1)(c + 1) = -VAL
            moduleMatrix(r + 1)(c + 2) = VAL

            moduleMatrix(r + 2)(c - 2) = VAL
            moduleMatrix(r + 2)(c - 1) = VAL
            moduleMatrix(r + 2)(c + 0) = VAL
            moduleMatrix(r + 2)(c + 1) = VAL
            moduleMatrix(r + 2)(c + 2) = VAL
Continue:
        Next
    Next
End Sub
