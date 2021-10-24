VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_buf() As Variant

Private Sub Class_Initialize()
    m_buf = Array()
End Sub

Public Property Get Count() As Long
    Count = UBound(m_buf) + 1
End Property

Public Sub Add(ByVal arg As Variant)
    ReDim Preserve m_buf(UBound(m_buf) + 1)

    If (VarType(arg) And VbVarType.vbObject) = VbVarType.vbObject Then
        Set m_buf(UBound(m_buf)) = arg
    Else
        m_buf(UBound(m_buf)) = arg
    End If
End Sub

Public Function Items() As Variant()
    Items = m_buf
End Function
