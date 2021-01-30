VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ByteEncoder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IQRCodeEncoder

Private m_data()        As Byte
Private m_charCounter   As Long
Private m_bitCounter    As Long

Private m_encAlpha  As IQRCodeEncoder
Private m_encKanji  As IQRCodeEncoder

Private m_textStream  As QRCodeLib.TextStream

Public Sub Init(ByVal charsetName As String)
    Call IQRCodeEncoder_Init(charsetName)
End Sub

Private Sub IQRCodeEncoder_Init(ByVal charsetName As String)
    Set m_textStream = NewTextStream(charsetName)
    Set m_encAlpha = New AlphanumericEncoder
    
    If Encoding.IsJP(charsetName) Then
        Set m_encKanji = NewKanjiEncoder(charsetName)
    Else
        ' NOP
    End If
End Sub

Private Property Get IQRCodeEncoder_BitCount() As Long
    IQRCodeEncoder_BitCount = m_bitCounter
End Property

Private Property Get IQRCodeEncoder_CharCount() As Long
    IQRCodeEncoder_CharCount = m_charCounter
End Property

Private Property Get IQRCodeEncoder_EncodingMode() As EncodingMode
    IQRCodeEncoder_EncodingMode = EncodingMode.EIGHT_BIT_BYTE
End Property

Private Property Get IQRCodeEncoder_ModeIndicator() As Long
    IQRCodeEncoder_ModeIndicator = ModeIndicator.BYTE_VALUE
End Property

Private Function IQRCodeEncoder_Append(ByVal c As String) As Long
    Dim charBytes() As Byte
    charBytes = m_textStream.GetBytes(c)

    Dim v   As Variant

    For Each v In charBytes
        If m_charCounter = 0 Then
            ReDim m_data(0)
        Else
            ReDim Preserve m_data(UBound(m_data) + 1)
        End If

        m_data(UBound(m_data)) = v
    Next
    
    Dim ret As Long
    ret = 8 * (UBound(charBytes) + 1)
    m_bitCounter = m_bitCounter + ret
    m_charCounter = m_charCounter + (UBound(charBytes) + 1)
    
    IQRCodeEncoder_Append = ret
End Function

Private Function IQRCodeEncoder_GetCodewordBitLength(ByVal c As String) As Long
    Dim charBytes() As Byte
    charBytes = m_textStream.GetBytes(c)

    IQRCodeEncoder_GetCodewordBitLength = (UBound(charBytes) + 1) * 8
End Function

Private Function IQRCodeEncoder_GetBytes() As Byte()
    IQRCodeEncoder_GetBytes = m_data
End Function

Private Function IQRCodeEncoder_InSubset(ByVal c As String) As Boolean
    IQRCodeEncoder_InSubset = True
End Function

Private Function IQRCodeEncoder_InExclusiveSubset(ByVal c As String) As Boolean
    If m_encAlpha.InSubset(c) Then
        IQRCodeEncoder_InExclusiveSubset = False
        Exit Function
    End If

    If Not (m_encKanji Is Nothing) Then
        If m_encKanji.InSubset(c) Then
            IQRCodeEncoder_InExclusiveSubset = False
            Exit Function
        End If
    End If
    
    If IQRCodeEncoder_InSubset(c) Then
        IQRCodeEncoder_InExclusiveSubset = True
        Exit Function
    End If

    IQRCodeEncoder_InExclusiveSubset = False
End Function
