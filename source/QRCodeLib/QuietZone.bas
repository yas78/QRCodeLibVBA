Attribute VB_Name = "QuietZone"
Option Private Module
Option Explicit

Public Const MIN_WIDTH As Long = 4

Private m_width As Long

Public Property Get Width() As Long
    If m_width = 0 Then
        m_width = MIN_WIDTH
    End If
    Width = m_width
End Property

Public Property Let Width(ByVal Value As Long)
    m_width = Value
End Property

Public Function Place(ByRef moduleMatrix() As Variant) As Variant()
    Dim sz As Long
    sz = UBound(moduleMatrix) + Width * 2

    Dim ret() As Variant
    ReDim ret(sz)

    Dim rowArray() As Long

    Dim i As Long
    For i = 0 To sz
        ReDim rowArray(sz)
        ret(i) = rowArray
    Next

    Dim r As Long
    Dim c As Long
    For r = 0 To UBound(moduleMatrix)
        For c = 0 To UBound(moduleMatrix(r))
            ret(r + Width)(c + Width) = moduleMatrix(r)(c)
        Next
    Next

    Place = ret
End Function
