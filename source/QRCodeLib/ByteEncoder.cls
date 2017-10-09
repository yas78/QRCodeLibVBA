VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ByteEncoder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------
' バイトモードエンコーダー
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

    IQRCodeEncoder_EncodingMode = EncodingMode.EIGHT_BIT_BYTE
    
End Property

'----------------------------------------------------------------------------------------
' (概要)
'  モード指示子を取得します。
'----------------------------------------------------------------------------------------
Private Property Get IQRCodeEncoder_ModeIndicator() As Long

    IQRCodeEncoder_ModeIndicator = ModeIndicator.BYTE_VALUE
    
End Property

'----------------------------------------------------------------------------------------
' (概要)
'  現在のモードシーケンスに文字を追加し、追加に要したビット数を返します。
'----------------------------------------------------------------------------------------
Private Function IQRCodeEncoder_Append(ByVal c As String) As Long

#If [DEBUG] Then
    Debug.Assert Len(c) = 1
#End If

    If m_charCounter = 0 Then
        ReDim m_data(0)
    Else
        ReDim Preserve m_data(UBound(m_data) + 1)
    End If
    
    Dim wd As Long
    wd = Asc(c) And &HFFFF&
    
    m_data(UBound(m_data)) = wd
    
    Dim ret As Long
    
    If wd > &HFF Then
        m_charCounter = m_charCounter + 2
        ret = 16
    Else
        m_charCounter = m_charCounter + 1
        ret = 8
    End If
    
    m_bitCounter = m_bitCounter + ret
    
    IQRCodeEncoder_Append = ret

End Function

'----------------------------------------------------------------------------------------
' (概要)
'  指定の文字をエンコードしたコード語のビット数を返します。
'----------------------------------------------------------------------------------------
Private Function IQRCodeEncoder_GetCodewordBitLength(ByVal c As String) As Long
    
#If [DEBUG] Then
    Debug.Assert Len(c) = 1
#End If

    Dim code    As Long
    code = Asc(c) And &HFFFF&
    
    IQRCodeEncoder_GetCodewordBitLength = IIf(code > &HFF, 16, 8)

End Function

'----------------------------------------------------------------------------------------
' (概要)
'  エンコードされたバイトデータを返します。
'----------------------------------------------------------------------------------------
Private Function IQRCodeEncoder_GetBytes() As Byte()
    
    Dim byteCount   As Long
    Dim ret()       As Byte
    Dim i           As Long
    
    Dim idx As Long
    
    Dim msb As Long
    Dim lsb As Long
    
    byteCount = m_bitCounter \ 8 + IIf((m_bitCounter Mod 8) > 0, 1, 0)
    ReDim ret(byteCount - 1)
    
    For i = 0 To UBound(m_data)
        msb = (m_data(i) And &HFF00&) \ 2 ^ 8
        lsb = m_data(i) And &HFF
        
        ret(idx) = IIf(m_data(i) > &HFF, msb, lsb)
        idx = idx + 1
        
        If m_data(i) > &HFF Then
            ret(idx) = lsb
            idx = idx + 1
        End If
    Next

    IQRCodeEncoder_GetBytes = ret
    
End Function

'----------------------------------------------------------------------------------------
' (概要)
'  指定の文字が符号化モードに適合する場合は True を返します。
'----------------------------------------------------------------------------------------
Private Function IQRCodeEncoder_IsInSubset(ByVal c As String) As Boolean

#If [DEBUG] Then
    Debug.Assert Len(c) = 1
#End If

    IQRCodeEncoder_IsInSubset = True
    
End Function

'----------------------------------------------------------------------------------------
' (概要)
'  指定の文字が符号化モードの排他的部分集合の場合は True を返します。
'----------------------------------------------------------------------------------------
Private Function IQRCodeEncoder_IsInExclusiveSubset(ByVal c As String) As Boolean

#If [DEBUG] Then
    Debug.Assert Len(c) = 1
#End If

    Dim ret As Boolean
    
    Dim code As Long
    code = Asc(c) And &HFFFF&
    
    Select Case code
        Case &H20&, &H24&, &H25&, &H2A&, &H2B&, &H2D& To &H3A&, &H41& To &H5A&
            ret = False
            
        Case &H8140& To &H9FFC&
            ret = False
    
        Case &HE040& To &HEBBF&
            ret = False
            
        Case Else
            ret = True
        
    End Select
    
    IQRCodeEncoder_IsInExclusiveSubset = ret
    
End Function

