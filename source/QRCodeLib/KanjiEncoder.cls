VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KanjiEncoder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------
' 漢字モードエンコーダー
'----------------------------------------------------------------------------------------
Option Explicit

Implements IQRCodeEncoder

Private m_data()        As Long
Private m_charCounter   As Long
Private m_bitCounter    As Long

'----------------------------------------------------------------------------------------
' (概要)
'  ビット数を取得します。
'----------------------------------------------------------------------------------------
Private Property Get IQRCodeEncoder_BitCount() As Long

    IQRCodeEncoder_BitCount = m_bitCounter
    
End Property

'----------------------------------------------------------------------------------------
' (概要)
'  文字数を取得します。
'----------------------------------------------------------------------------------------
Private Property Get IQRCodeEncoder_CharCount() As Long

    IQRCodeEncoder_CharCount = m_charCounter
    
End Property

'----------------------------------------------------------------------------------------
' (概要)
'  符号化モードを表す EncodingMode 列挙体の値を取得します。
'----------------------------------------------------------------------------------------
Private Property Get IQRCodeEncoder_EncodingMode() As EncodingMode

    IQRCodeEncoder_EncodingMode = EncodingMode.KANJI
    
End Property

'----------------------------------------------------------------------------------------
' (概要)
'  モード指示子を取得します。
'----------------------------------------------------------------------------------------
Private Property Get IQRCodeEncoder_ModeIndicator() As Long

    IQRCodeEncoder_ModeIndicator = ModeIndicator.KANJI_VALUE
    
End Property

'----------------------------------------------------------------------------------------
' (概要)
'  現在のモードシーケンスに文字を追加し、追加に要したビット数を返します。
'----------------------------------------------------------------------------------------
Private Function IQRCodeEncoder_Append(ByVal c As String) As Long
    
#If [DEBUG] Then
    Debug.Assert Len(c) = 1
    Debug.Assert IQRCodeEncoder_IsInSubset(c)
#End If

    Dim wd  As Long
    wd = Asc(c) And &HFFFF&
    
    Select Case wd
        Case &H8140& To &H9FFC&
            wd = wd - &H8140&
        Case &HE040& To &HEBBF&
            wd = wd - &HC140&
        Case Else
            Err.Raise 5
    End Select

    wd = ((wd \ 2 ^ 8) * &HC0&) + (wd And &HFF&)
    If m_charCounter = 0 Then
        ReDim m_data(0)
    Else
        ReDim Preserve m_data(UBound(m_data) + 1)
    End If
    
    m_data(UBound(m_data)) = wd
    
    m_charCounter = m_charCounter + 1
    m_bitCounter = m_bitCounter + 13
    
    IQRCodeEncoder_Append = 13
    
End Function

'----------------------------------------------------------------------------------------
' (概要)
'  指定の文字をエンコードしたコード語のビット数を返します。
'----------------------------------------------------------------------------------------
Private Function IQRCodeEncoder_GetCodewordBitLength(ByVal c As String) As Long

#If [DEBUG] Then
    Debug.Assert Len(c) = 1
    Debug.Assert IQRCodeEncoder_IsInSubset(c)
#End If

    IQRCodeEncoder_GetCodewordBitLength = 13
     
End Function

'----------------------------------------------------------------------------------------
' (概要)
'  エンコードされたバイトデータを返します。
'----------------------------------------------------------------------------------------
Private Function IQRCodeEncoder_GetBytes() As Byte()

    Dim bs As BitSequence
    Set bs = New BitSequence

    Dim i As Long
    
    For i = 0 To UBound(m_data)
        Call bs.Append(m_data(i), 13)
    Next

    IQRCodeEncoder_GetBytes = bs.GetBytes()
    
End Function

'----------------------------------------------------------------------------------------
' (概要)
'  指定の文字が符号化モードに適合する場合は True を返します。
'----------------------------------------------------------------------------------------
Private Function IQRCodeEncoder_IsInSubset(ByVal c As String) As Boolean

#If [DEBUG] Then
    Debug.Assert Len(c) = 1
#End If

    Dim code As Long
    code = Asc(c) And &HFFFF&
    
    Dim lsb As Long
    lsb = code And &HFF&
    
    If code >= &H8140& And code <= &H9FFC& Or _
       code >= &HE040& And code <= &HEBBF& Then
       
        IQRCodeEncoder_IsInSubset = lsb >= &H40& And _
                                    lsb <= &HFC& And _
                                    lsb <> &H7F&
    Else
        IQRCodeEncoder_IsInSubset = False
    End If
    
End Function

'----------------------------------------------------------------------------------------
' (概要)
'  指定の文字が符号化モードの排他的部分集合の場合は True を返します。
'----------------------------------------------------------------------------------------
Private Function IQRCodeEncoder_IsInExclusiveSubset(ByVal c As String) As Boolean

    IQRCodeEncoder_IsInExclusiveSubset = IQRCodeEncoder_IsInSubset(c)
    
End Function

