VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AlphanumericEncoder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
' 英数字モードエンコーダー
'------------------------------------------------------------------------------
Option Explicit

Implements IQRCodeEncoder

Private m_data()        As Long
Private m_charCounter   As Long
Private m_bitCounter    As Long

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
    
    IQRCodeEncoder_EncodingMode = EncodingMode.ALPHA_NUMERIC

End Property

'------------------------------------------------------------------------------
' (概要)
'  モード指示子を取得します。
'------------------------------------------------------------------------------
Private Property Get IQRCodeEncoder_ModeIndicator() As Long
    
    IQRCodeEncoder_ModeIndicator = ModeIndicator.ALPAHNUMERIC_VALUE

End Property

'------------------------------------------------------------------------------
' (概要)
'  現在のモードシーケンスに文字を追加し、追加に要したビット数を返します。
'------------------------------------------------------------------------------
Private Function IQRCodeEncoder_Append(ByVal c As String) As Long

#If [DEBUG] Then
    Debug.Assert Len(c) = 1
    Debug.Assert IQRCodeEncoder_IsInSubset(c)
#End If

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

'------------------------------------------------------------------------------
' (概要)
'  指定の文字をエンコードしたコード語のビット数を返します。
'------------------------------------------------------------------------------
Private Function IQRCodeEncoder_GetCodewordBitLength(ByVal c As String) As Long

#If [DEBUG] Then
    Debug.Assert Len(c) = 1
    Debug.Assert IQRCodeEncoder_IsInSubset(c)
#End If

    If m_charCounter Mod 2 = 0 Then
        IQRCodeEncoder_GetCodewordBitLength = 6
    Else
        IQRCodeEncoder_GetCodewordBitLength = 5
    End If
    
End Function

'------------------------------------------------------------------------------
' (概要)
'  エンコードされたバイトデータを返します。
'------------------------------------------------------------------------------
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

'------------------------------------------------------------------------------
' (概要)
'  指定の文字が符号化モードに適合する場合は True を返します。
'------------------------------------------------------------------------------
Private Function IQRCodeEncoder_IsInSubset(ByVal c As String) As Boolean

#If [DEBUG] Then
    Debug.Assert Len(c) = 1
#End If

    Dim ret As Boolean
    
    Select Case Asc(c)
        ' A - Z
        Case 65 To 90
            ret = True
        ' 0 - 9
        Case 48 To 57
            ret = True
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
    
        Case Else
            ret = False
            
    End Select

    IQRCodeEncoder_IsInSubset = ret

End Function

'------------------------------------------------------------------------------
' (概要)
'  指定の文字が符号化モードの排他的部分集合の場合は True を返します。
'------------------------------------------------------------------------------
Private Function IQRCodeEncoder_IsInExclusiveSubset(ByVal c As String) As Boolean

#If [DEBUG] Then
    Debug.Assert Len(c) = 1
#End If

    Dim ret As Boolean
    
    Select Case Asc(c)
        ' A - Z
        Case 65 To 90
            ret = True
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
    
        Case Else
            ret = False
        
    End Select
    
    IQRCodeEncoder_IsInExclusiveSubset = ret

End Function

'------------------------------------------------------------------------------
' (概要)
'  指定した文字の英数字モードにおける文字コードを返します。該当する文字コードが存在しない
'  場合は -1 を返します。
'------------------------------------------------------------------------------
Private Function ConvertCharCode(ByVal c As String) As Long

#If [DEBUG] Then
    Debug.Assert Len(c) = 1
#End If

    Dim code As Long
    code = Asc(c)

    Select Case code
        ' A - Z
        Case 65 To 90
            ConvertCharCode = code - 55
        ' 0 - 9
        Case 48 To 57
            ConvertCharCode = code - 48
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
        ' :
        Case 58
            ConvertCharCode = 44
    
        Case Else
            ConvertCharCode = -1
            
    End Select

End Function

