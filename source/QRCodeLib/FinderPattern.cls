VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FinderPattern"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const VAL As Long = Values.FINDER_PTN

Private m_finderPattern() As Variant

Private Sub Class_Initialize()
    m_finderPattern = Array( _
        Array(VAL, VAL, VAL, VAL, VAL, VAL, VAL), _
        Array(VAL, -VAL, -VAL, -VAL, -VAL, -VAL, VAL), _
        Array(VAL, -VAL, VAL, VAL, VAL, -VAL, VAL), _
        Array(VAL, -VAL, VAL, VAL, VAL, -VAL, VAL), _
        Array(VAL, -VAL, VAL, VAL, VAL, -VAL, VAL), _
        Array(VAL, -VAL, -VAL, -VAL, -VAL, -VAL, VAL), _
        Array(VAL, VAL, VAL, VAL, VAL, VAL, VAL) _
    )
End Sub

Public Sub Place(ByRef moduleMatrix() As Variant)
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
