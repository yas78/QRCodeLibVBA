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

Private m_items() As Variant

Private Sub Class_Initialize()
    m_items = Array()
End Sub

Public Property Get Count() As Long
    Count = UBound(m_items) + 1
End Property

Public Sub Add(ByVal arg As Variant)
    ReDim Preserve m_items(UBound(m_items) + 1)

    If (VarType(arg) And VbVarType.vbObject) = VbVarType.vbObject Then
        Set m_items(UBound(m_items)) = arg
    Else
        m_items(UBound(m_items)) = arg
    End If
End Sub

Public Function Item(ByVal Index As Long) As Variant
Attribute Item.VB_UserMemId = 0
' Default Member [Attribute Item.VB_UserMemId = 0]
    If (VarType(m_items(Index)) And VbVarType.vbObject) = VbVarType.vbObject Then
        Set Item = m_items(Index)
    Else
        Item = m_items(Index)
    End If
End Function

Public Function Items() As Variant()
    Items = m_items
End Function
