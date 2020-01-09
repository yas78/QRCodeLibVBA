VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NumericEncoder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IQRCodeEncoder

Private m_data()        As Long
Private m_charCounter   As Long
Private m_bitCounter    As Long

Private Property Get IQRCodeEncoder_BitCount() As Long
    IQRCodeEncoder_BitCount = m_bitCounter
End Property

Private Property Get IQRCodeEncoder_CharCount() As Long
    IQRCodeEncoder_CharCount = m_charCounter
End Property

Private Property Get IQRCodeEncoder_EncodingMode() As EncodingMode
    IQRCodeEncoder_EncodingMode = EncodingMode.NUMERIC
End Property

Private Property Get IQRCodeEncoder_ModeIndicator() As Long
    IQRCodeEncoder_ModeIndicator = ModeIndicator.NUMERIC_VALUE
End Property

Private Function IQRCodeEncoder_Append(ByVal c As String) As Long
    Dim ret As Long

    If m_charCounter Mod 3 = 0 Then
        If m_charCounter = 0 Then
            ReDim m_data(0)
        Else
            ReDim Preserve m_data(UBound(m_data) + 1)
        End If

        m_data(UBound(m_data)) = CLng(c)
        ret = 4

    Else
        m_data(UBound(m_data)) = m_data(UBound(m_data)) * 10 + CLng(c)
        ret = 3

    End If

    m_charCounter = m_charCounter + 1
    m_bitCounter = m_bitCounter + ret

    IQRCodeEncoder_Append = ret
End Function

Private Function IQRCodeEncoder_GetCodewordBitLength(ByVal c As String) As Long
    If m_charCounter Mod 3 = 0 Then
        IQRCodeEncoder_GetCodewordBitLength = 4
    Else
        IQRCodeEncoder_GetCodewordBitLength = 3
    End If
End Function

Private Function IQRCodeEncoder_GetBytes() As Byte()
    Dim bs As BitSequence
    Set bs = New BitSequence

    Dim i As Long

    For i = 0 To UBound(m_data) - 1
        Call bs.Append(m_data(i), 10)
    Next

    Select Case m_charCounter Mod 3
        Case 1
            Call bs.Append(m_data(UBound(m_data)), 4)
        Case 2
            Call bs.Append(m_data(UBound(m_data)), 7)
        Case Else
            Call bs.Append(m_data(UBound(m_data)), 10)
    End Select

    IQRCodeEncoder_GetBytes = bs.GetBytes()
End Function

Private Function IQRCodeEncoder_InSubset(ByVal c As String) As Boolean
    IQRCodeEncoder_InSubset = c Like "[0-9]"
End Function

Private Function IQRCodeEncoder_InExclusiveSubset(ByVal c As String) As Boolean
    IQRCodeEncoder_InExclusiveSubset = IQRCodeEncoder_InSubset(c)
End Function
