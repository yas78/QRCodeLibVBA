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

#If VBA7 Then
    Private Declare PtrSafe Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal pDest As LongPtr, ByVal pSrc As LongPtr, ByVal sz As Long)
#Else
    Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal pDest As Long, ByVal pSrc As Long, ByVal sz As Long)
#End If

Private m_buffer()   As Byte
Private m_bitCounter As Long
Private m_space      As Long
Private m_index      As Long

Private Sub Class_Initialize()
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
End Sub

Public Sub Append(ByVal data As Long, ByVal bitLength As Long)
    Dim remainingLength As Long
    remainingLength = bitLength

    Dim remainingData As Long
    remainingData = data

    Dim temp As Byte

    Do While remainingLength > 0
        If m_space = 0 Then
            m_space = 8
            m_index = m_index + 1
            ReDim Preserve m_buffer(m_index)
        End If

        temp = m_buffer(m_index)

        If m_space < remainingLength Then
            temp = CByte(temp Or remainingData \ (2 ^ (remainingLength - m_space)))

            remainingData = remainingData And ((2 ^ (remainingLength - m_space)) - 1)

            m_bitCounter = m_bitCounter + m_space
            remainingLength = remainingLength - m_space
            m_space = 0
        Else
            temp = CByte(temp Or remainingData * (2 ^ (m_space - remainingLength)))

            m_bitCounter = m_bitCounter + remainingLength
            m_space = m_space - remainingLength
            remainingLength = 0
        End If

        m_buffer(m_index) = temp
    Loop
End Sub

Public Function GetBytes() As Byte()
    If m_index < 0 Then Call Err.Raise(51)

    Dim ret() As Byte
    ReDim ret(m_index)

    Call MoveMemory(VarPtr(ret(0)), VarPtr(m_buffer(0)), m_index + 1)
    GetBytes = ret
End Function
