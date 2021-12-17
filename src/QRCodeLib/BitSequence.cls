VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BitSequence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum PackingOrder
    MSBFirst
    LSBFirst
End Enum

Private m_buffer()   As Byte
Private m_bitCounter As Long
Private m_space      As Long
Private m_index      As Long
Private m_pos        As Long
Private m_order      As PackingOrder

Private Sub Class_Initialize()
    Call Init(PackingOrder.MSBFirst)
End Sub

Public Sub Init(ByVal order As PackingOrder)
    m_order = order
    Call Clear
End Sub

Public Property Get Length() As Long
    Length = m_bitCounter
End Property

Public Sub Clear()
    Erase m_buffer
    m_index = -1
    m_bitCounter = 0
    m_space = 0

    Select Case m_order
        Case PackingOrder.MSBFirst
            m_pos = -1
        Case PackingOrder.LSBFirst
            m_pos = 8
        Case Else
            Call Err.Raise(51)
    End Select
End Sub

Public Sub Append(ByVal data As Long, ByVal sz As Long)
    Dim s As Long
    Dim e As Long
    Dim step As Long

    Select Case m_order
        Case PackingOrder.MSBFirst
            s = sz - 1
            e = 0
            step = -1
        Case PackingOrder.LSBFirst
            s = 0
            e = sz - 1
            step = 1
        Case Else
            Call Err.Raise(51)
    End Select

    Dim i As Long
    For i = s To e Step step
        If m_space = 0 Then
            m_space = 8
            m_index = m_index + 1
            ReDim Preserve m_buffer(m_index)

            Select Case m_order
                Case PackingOrder.MSBFirst
                    m_pos = 7
                Case PackingOrder.LSBFirst
                    m_pos = 0
                Case Else
                    Call Err.Raise(51)
            End Select
        End If

        m_buffer(m_index) = m_buffer(m_index) Or ((data And 2 ^ i) \ (2 ^ i)) * (2 ^ m_pos)
        m_space = m_space - 1
        m_pos = m_pos + step
    Next

    m_bitCounter = m_bitCounter + sz
End Sub

Public Function GetBytes() As Byte()
    If m_index < 0 Then Call Err.Raise(51)
    GetBytes = m_buffer
End Function
