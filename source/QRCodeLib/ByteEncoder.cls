VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ByteEncoder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
' バイトモードエンコーダー
'------------------------------------------------------------------------------
Option Explicit

Implements IQRCodeEncoder

Private m_data()        As Byte
Private m_charCounter   As Long
Private m_bitCounter    As Long

Private m_encAlpha  As IQRCodeEncoder
Private m_encKanji  As IQRCodeEncoder

Private m_textEncoding  As TextEncoding

Private Sub Class_Initialize()

    Set m_encAlpha = New AlphanumericEncoder
    Set m_encKanji = New KanjiEncoder
    Set m_textEncoding = New TextEncoding

End Sub

Public Sub Initialize(ByVal charsetName As String)

    Call m_textEncoding.Init(charsetName)

End Sub

'------------------------------------------------------------------------------
' (概要)
'  ビット数を取得します。
'------------------------------------------------------------------------------
Private Property Get IQRCodeEncoder_BitCount() As Long

    IQRCodeEncoder_BitCount = m_bitCounter

End Property

'------------------------------------------------------------------------------
' (概要)
'  文字数を取得します。
'------------------------------------------------------------------------------
Private Property Get IQRCodeEncoder_CharCount() As Long

    IQRCodeEncoder_CharCount = m_charCounter

End Property

'------------------------------------------------------------------------------
' (概要)
'  符号化モードを表す EncodingMode 列挙体の値を取得します。
'------------------------------------------------------------------------------
Private Property Get IQRCodeEncoder_EncodingMode() As EncodingMode

    IQRCodeEncoder_EncodingMode = EncodingMode.EIGHT_BIT_BYTE

End Property

'------------------------------------------------------------------------------
' (概要)
'  モード指示子を取得します。
'------------------------------------------------------------------------------
Private Property Get IQRCodeEncoder_ModeIndicator() As Long

    IQRCodeEncoder_ModeIndicator = ModeIndicator.BYTE_VALUE

End Property

'------------------------------------------------------------------------------
' (概要)
'  現在のモードシーケンスに文字を追加し、追加に要したビット数を返します。
'------------------------------------------------------------------------------
Private Function IQRCodeEncoder_Append(ByVal c As String) As Long

    Dim charBytes() As Byte
    charBytes = m_textEncoding.GetBytes(c)

    Dim ret As Long
    Dim v   As Variant

    For Each v In charBytes
        If m_charCounter = 0 Then
            ReDim m_data(0)
        Else
            ReDim Preserve m_data(UBound(m_data) + 1)
        End If

        m_data(UBound(m_data)) = v
        m_charCounter = m_charCounter + 1
        m_bitCounter = m_bitCounter + 8
        ret = ret + 8
    Next

    IQRCodeEncoder_Append = ret

End Function

'------------------------------------------------------------------------------
' (概要)
'  指定の文字をエンコードしたコード語のビット数を返します。
'------------------------------------------------------------------------------
Private Function IQRCodeEncoder_GetCodewordBitLength(ByVal c As String) As Long

    Dim charBytes() As Byte
    charBytes = m_textEncoding.GetBytes(c)

    IQRCodeEncoder_GetCodewordBitLength = (UBound(charBytes) + 1) * 8

End Function

'------------------------------------------------------------------------------
' (概要)
'  エンコードされたバイトデータを返します。
'------------------------------------------------------------------------------
Private Function IQRCodeEncoder_GetBytes() As Byte()

    IQRCodeEncoder_GetBytes = m_data

End Function

'------------------------------------------------------------------------------
' (概要)
'  指定の文字が符号化モードに適合する場合は True を返します。
'------------------------------------------------------------------------------
Private Function IQRCodeEncoder_InSubset(ByVal c As String) As Boolean

    IQRCodeEncoder_InSubset = True

End Function

'------------------------------------------------------------------------------
' (概要)
'  指定の文字が符号化モードの排他的部分集合の場合は True を返します。
'------------------------------------------------------------------------------
Private Function IQRCodeEncoder_InExclusiveSubset(ByVal c As String) As Boolean

    If m_encAlpha.InSubset(c) Then
        IQRCodeEncoder_InExclusiveSubset = False
        Exit Function
    End If

    If m_encKanji.InSubset(c) Then
        IQRCodeEncoder_InExclusiveSubset = False
        Exit Function
    End If

    If IQRCodeEncoder_InSubset(c) Then
        IQRCodeEncoder_InExclusiveSubset = True
        Exit Function
    End If

    IQRCodeEncoder_InExclusiveSubset = False

End Function
