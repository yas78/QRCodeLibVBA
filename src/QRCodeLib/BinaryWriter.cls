VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BinaryWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_length   As Long
Private m_buffer() As Byte
Private m_capacity As Long
Private m_position As Long

Private Sub Class_Initialize()
    Call Init
End Sub

Public Sub Init()
    m_length = 0
    ReDim m_buffer(&HFF)
    m_capacity = UBound(m_buffer) + 1
    m_position = 0
End Sub

Public Property Get Length() As Long
    Length = m_length
End Property

Public Property Get Capacity() As Long
    Capacity = m_capacity
End Property

Public Property Get Position() As Long
    Position = m_position
End Property
Public Property Let Position(ByVal Value As Long)
    If Value < 0 Then Call Err.Raise(5)
    m_position = Value
End Property

Private Sub WriteByte(ByVal Value As Byte)
    If (m_length + 1) > m_capacity Then
        ReDim Preserve m_buffer(m_capacity * 2 - 1)
        m_capacity = UBound(m_buffer) + 1
    End If

    m_length = m_length + 1
    m_buffer(m_position) = Value
    m_position = m_position + 1
End Sub

Public Sub WriteAny(ByVal arg As Variant, Optional ByVal reverse As Boolean = False)
    Dim t As VbVarType
    t = VarType(arg)

    Dim i As Long
    Dim j As Long
    Dim bytes() As Byte

    If (t And VbVarType.vbArray) = VbVarType.vbArray Then
        Select Case t - VbVarType.vbArray
            Case VbVarType.vbByte
                For i = 0 To UBound(arg)
                    Call WriteByte(arg(i))
                Next
            Case VbVarType.vbInteger, VbVarType.vbLong
                For i = 0 To UBound(arg)
                    bytes = BitConverter.GetBytes(arg(i), reverse)
                    For j = 0 To UBound(bytes)
                        Call WriteByte(bytes(j))
                    Next
                Next
#If Win64 Then
            Case VbVarType.vbLongLong
                For i = 0 To UBound(arg)
                    bytes = BitConverter.GetBytes(arg(i), reverse)
                    For j = 0 To UBound(bytes)
                        Call WriteByte(bytes(j))
                    Next
                Next
#End If
            Case Else
                Call Err.Raise(5)
        End Select
    Else
        Select Case VarType(arg)
            Case VbVarType.vbByte
                Call WriteByte(arg)
            Case VbVarType.vbInteger, VbVarType.vbLong
                bytes = BitConverter.GetBytes(arg, reverse)
                For i = 0 To UBound(bytes)
                    Call WriteByte(bytes(i))
                Next
#If Win64 Then
            Case VbVarType.vbLongLong
                bytes = BitConverter.GetBytes(arg, reverse)
                For i = 0 To UBound(bytes)
                    Call WriteByte(bytes(i))
                Next
#End If
            Case Else
                Call Err.Raise(5)
        End Select
    End If
End Sub

Public Function Flush2() As Byte()
    ReDim Preserve m_buffer(m_length - 1)
    Flush2 = m_buffer
    Call Init
End Function

Public Sub Flush(ByRef buffer() As Byte)
    ReDim buffer(m_length - 1)

    Dim i As Long
    For i = 0 To m_length - 1
        buffer(i) = m_buffer(i)
    Next

    Call Init
End Sub
