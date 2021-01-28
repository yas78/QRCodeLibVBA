VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KanjiEncoder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IQRCodeEncoder

Private m_data()        As Long
Private m_charCounter   As Long
Private m_bitCounter    As Long

Private m_alphaNumericEncoder As IQRCodeEncoder

Private m_textStream  As QRCodeLib.TextStream

Public Sub Init(ByVal charsetName As String)
    Call IQRCodeEncoder_Init(charsetName)
End Sub

Private Sub IQRCodeEncoder_Init(ByVal charsetName As String)
    Set m_alphaNumericEncoder = New AlphanumericEncoder
    Set m_textStream = NewTextStream(charsetName)
End Sub

Private Property Get IQRCodeEncoder_BitCount() As Long
    IQRCodeEncoder_BitCount = m_bitCounter
End Property

Private Property Get IQRCodeEncoder_CharCount() As Long
    IQRCodeEncoder_CharCount = m_charCounter
End Property

Private Property Get IQRCodeEncoder_EncodingMode() As EncodingMode
    IQRCodeEncoder_EncodingMode = EncodingMode.KANJI
End Property

Private Property Get IQRCodeEncoder_ModeIndicator() As Long
    IQRCodeEncoder_ModeIndicator = ModeIndicator.KANJI_VALUE
End Property

Private Function IQRCodeEncoder_Append(ByVal c As String) As Long
    Dim charBytes() As Byte
    charBytes = m_textStream.GetBytes(c)

    Dim wd As Long
    wd = (CLng(charBytes(0)) * 2 ^ 8) Or CLng(charBytes(1))

    Select Case wd
        Case &H8140& To &H9FFC&
            wd = wd - &H8140&
        Case &HE040& To &HEBBF&
            wd = wd - &HC140&
        Case Else
            Call Err.Raise(5)
    End Select

    wd = ((wd \ 2 ^ 8) * &HC0&) + (wd And &HFF&)
    If m_charCounter = 0 Then
        ReDim m_data(0)
    Else
        ReDim Preserve m_data(UBound(m_data) + 1)
    End If

    m_data(UBound(m_data)) = wd
    
    Dim ret As Long
    ret = IQRCodeEncoder_GetCodewordBitLength(c)
    m_bitCounter = m_bitCounter + ret
    m_charCounter = m_charCounter + 1

    IQRCodeEncoder_Append = ret
End Function

Private Function IQRCodeEncoder_GetCodewordBitLength(ByVal c As String) As Long
    IQRCodeEncoder_GetCodewordBitLength = 13
End Function

Private Function IQRCodeEncoder_GetBytes() As Byte()
    Dim bs As New BitSequence

    Dim v As Variant

    For Each v In m_data
        Call bs.Append(v, 13)
    Next

    IQRCodeEncoder_GetBytes = bs.GetBytes()
End Function

Private Function IQRCodeEncoder_InSubset(ByVal c As String) As Boolean
    Dim charBytes() As Byte
    charBytes = m_textStream.GetBytes(c)

    If UBound(charBytes) <> 1 Then
        IQRCodeEncoder_InSubset = False
        Exit Function
    End If

    Dim code As Long
    code = (CLng(charBytes(0)) * 2 ^ 8) Or CLng(charBytes(1))

    Dim lsb As Long
    lsb = code And &HFF&

    If &H8140& <= code And code <= &H9FFC& Or _
       &HE040& <= code And code <= &HEBBF& Then
        IQRCodeEncoder_InSubset = &H40& <= lsb And lsb <= &HFC& And _
                                  &H7F& <> lsb
        Exit Function
    End If
    
    IQRCodeEncoder_InSubset = False
End Function

Private Function IQRCodeEncoder_InExclusiveSubset(ByVal c As String) As Boolean
    If m_alphaNumericEncoder.InSubset(c) Then
        IQRCodeEncoder_InExclusiveSubset = False
        Exit Function
    End If

    IQRCodeEncoder_InExclusiveSubset = IQRCodeEncoder_InSubset(c)
End Function
