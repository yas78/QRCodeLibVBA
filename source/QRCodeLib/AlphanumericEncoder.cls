VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AlphanumericEncoder"
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
    IQRCodeEncoder_EncodingMode = EncodingMode.ALPHA_NUMERIC
End Property

Private Property Get IQRCodeEncoder_ModeIndicator() As Long
    IQRCodeEncoder_ModeIndicator = ModeIndicator.ALPAHNUMERIC_VALUE
End Property

Private Function IQRCodeEncoder_Append(ByVal c As String) As Long
    Dim wd As Long
    wd = ConvertCharCode(c)

    Dim ret As Long

    If m_charCounter Mod 2 = 0 Then
        If m_charCounter = 0 Then
            ReDim m_data(0)
        Else
            ReDim Preserve m_data(UBound(m_data) + 1)
        End If

        m_data(UBound(m_data)) = wd
        ret = 6
    Else
        m_data(UBound(m_data)) = m_data(UBound(m_data)) * 45
        m_data(UBound(m_data)) = m_data(UBound(m_data)) + wd
        ret = 5
    End If

    m_charCounter = m_charCounter + 1
    m_bitCounter = m_bitCounter + ret

    IQRCodeEncoder_Append = ret
End Function

Private Function IQRCodeEncoder_GetCodewordBitLength(ByVal c As String) As Long
    If m_charCounter Mod 2 = 0 Then
        IQRCodeEncoder_GetCodewordBitLength = 6
    Else
        IQRCodeEncoder_GetCodewordBitLength = 5
    End If
End Function

Private Function IQRCodeEncoder_GetBytes() As Byte()
    Dim bs As BitSequence
    Set bs = New BitSequence

    Dim bitLength As Long
    bitLength = 11

    Dim i As Long

    For i = 0 To UBound(m_data) - 1
        Call bs.Append(m_data(i), bitLength)
    Next

    If m_charCounter Mod 2 = 0 Then
        bitLength = 11
    Else
        bitLength = 6
    End If

    Call bs.Append(m_data(UBound(m_data)), bitLength)

    IQRCodeEncoder_GetBytes = bs.GetBytes()
End Function

Private Function IQRCodeEncoder_InSubset(ByVal c As String) As Boolean
    Dim ret As Boolean

    Select Case Asc(c)
        ' (Space)
        Case 32
            ret = True
        ' $ %
        Case 36, 37
            ret = True
        ' * +
        Case 42, 43
            ret = True
        ' - .
        Case 45, 46
            ret = True
        ' /
        Case 47
            ret = True
        ' 0 - 9
        Case 48 To 57
            ret = True
        ' :
        Case 58
            ret = True
        ' A - Z
        Case 65 To 90
            ret = True

        Case Else
            ret = False
    End Select

    IQRCodeEncoder_InSubset = ret
End Function

Private Function IQRCodeEncoder_InExclusiveSubset(ByVal c As String) As Boolean
    Dim ret As Boolean

    Select Case Asc(c)
        ' (Space)
        Case 32
            ret = True
        ' $ %
        Case 36, 37
            ret = True
        ' * +
        Case 42, 43
            ret = True
        ' - .
        Case 45, 46
            ret = True
        ' /
        Case 47
            ret = True
        ' :
        Case 58
            ret = True
        ' A - Z
        Case 65 To 90
            ret = True

        Case Else
            ret = False

    End Select

    IQRCodeEncoder_InExclusiveSubset = ret
End Function

Private Function ConvertCharCode(ByVal c As String) As Long
    Dim code As Long
    code = Asc(c)

    Select Case code
        ' (Space)
        Case 32
            ConvertCharCode = 36
        ' $ %
        Case 36, 37
            ConvertCharCode = code + 1
        ' * +
        Case 42, 43
            ConvertCharCode = code - 3
        ' - .
        Case 45, 46
            ConvertCharCode = code - 4
        ' /
        Case 47
            ConvertCharCode = 43
        ' 0 - 9
        Case 48 To 57
            ConvertCharCode = code - 48
        ' :
        Case 58
            ConvertCharCode = 44
        ' A - Z
        Case 65 To 90
            ConvertCharCode = code - 55

        Case Else
            ConvertCharCode = -1
    End Select
End Function
