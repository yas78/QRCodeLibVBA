VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Symbols"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_items  As Collection

Private m_minVersion              As Long
Private m_maxVersion              As Long
Private m_errorCorrectionLevel    As ErrorCorrectionLevel
Private m_structuredAppendAllowed As Boolean
Private m_byteModeCharsetName     As String

Private m_parity As Long

Private m_currSymbol As Symbol

Private m_encNum    As IQRCodeEncoder
Private m_encAlpha  As IQRCodeEncoder
Private m_encByte   As IQRCodeEncoder
Private m_encKanji  As IQRCodeEncoder

Private m_byteModeEncoding As TextEncoding
Private m_shiftJISEncoding As TextEncoding

Friend Sub Init(ByVal ecLevel As ErrorCorrectionLevel, _
                ByVal maxVer As Long, _
                ByVal allowStructuredAppend As Boolean, _
                ByVal byteModeCharsetName As String)
    If Not (MIN_VERSION <= maxVer And maxVer <= MAX_VERSION) Then
        Call Err.Raise(5)
    End If

    Set m_items = New Collection

    Set m_encNum = Encoder.Create(EncodingMode.NUMERIC)
    Set m_encAlpha = Encoder.Create(EncodingMode.ALPHA_NUMERIC)
    Set m_encByte = Encoder.Create(EncodingMode.EIGHT_BIT_BYTE, byteModeCharsetName)
    Set m_encKanji = Encoder.Create(EncodingMode.KANJI)

    m_minVersion = 1
    m_maxVersion = maxVer
    m_errorCorrectionLevel = ecLevel
    m_structuredAppendAllowed = allowStructuredAppend
    m_byteModeCharsetName = byteModeCharsetName

    m_parity = 0

    Set m_byteModeEncoding = New TextEncoding
    Call m_byteModeEncoding.Init(byteModeCharsetName)
    Set m_shiftJISEncoding = New TextEncoding
    Call m_shiftJISEncoding.Init("Shift_JIS")

    Set m_currSymbol = New Symbol
    Call m_currSymbol.Init(Me)
    Call m_items.Add(m_currSymbol)
End Sub

Public Property Get Item(ByVal idx As Long) As Symbol
Attribute Item.VB_UserMemId = 0
' Default Member [Attribute Item.VB_UserMemId = 0]
    Set Item = m_items(idx + 1)
End Property

Public Property Get Count() As Long
    Count = m_items.Count
End Property

Friend Property Get StructuredAppendAllowed() As Boolean
    StructuredAppendAllowed = m_structuredAppendAllowed
End Property

Friend Property Get Parity() As Long
    Parity = m_parity
End Property

Friend Property Get MinVersion() As Long
    MinVersion = m_minVersion
End Property
Friend Property Let MinVersion(ByVal Value As Long)
    m_minVersion = Value
End Property

Friend Property Get MaxVersion() As Long
    MaxVersion = m_maxVersion
End Property

Friend Property Get ErrorCorrectionLevel() As ErrorCorrectionLevel
    ErrorCorrectionLevel = m_errorCorrectionLevel
End Property

Friend Property Get ByteModeCharset() As String
    ByteModeCharset = m_byteModeCharsetName
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
' [Attribute NewEnum.VB_UserMemId = -4]
    Set NewEnum = m_items.[_NewEnum]
End Property

Private Function Add() As Symbol
    Set m_currSymbol = New Symbol
    Call m_currSymbol.Init(Me)
    Call m_items.Add(m_currSymbol)

    Set Add = m_currSymbol
End Function

Public Sub AppendText(ByVal s As String)
    If Len(s) = 0 Then Call Err.Raise(5)

    Dim oldMode As EncodingMode
    Dim newMode As EncodingMode

    Dim i As Long
    Dim c As String

    For i = 1 To Len(s)
        oldMode = m_currSymbol.CurrentEncodingMode

        Select Case oldMode
            Case EncodingMode.UNKNOWN
                newMode = SelectInitialMode(s, i)
            Case EncodingMode.NUMERIC
                newMode = SelectModeWhileInNumericMode(s, i)
            Case EncodingMode.ALPHA_NUMERIC
                newMode = SelectModeWhileInAlphanumericMode(s, i)
            Case EncodingMode.EIGHT_BIT_BYTE
                newMode = SelectModeWhileInByteMode(s, i)
            Case EncodingMode.KANJI
                newMode = SelectInitialMode(s, i)
            Case Else
                Call Err.Raise(51)
        End Select

        c = Mid$(s, i, 1)

        If newMode <> oldMode Then
            If Not m_currSymbol.TrySetEncodingMode(newMode, c) Then
                If Not m_structuredAppendAllowed Or m_items.Count = 16 Then
                    Call Err.Raise(6, Description:="String too long")
                End If

                Call Add
                newMode = SelectInitialMode(s, i)
                Call m_currSymbol.TrySetEncodingMode(newMode, c)
            End If
        End If

        If Not m_currSymbol.TryAppend(c) Then
            If Not m_structuredAppendAllowed Or m_items.Count = 16 Then
                Call Err.Raise(6, Description:="String too long")
            End If

            Call Add
            newMode = SelectInitialMode(s, i)
            Call m_currSymbol.TrySetEncodingMode(newMode, c)
            Call m_currSymbol.TryAppend(c)
        End If
    Next
End Sub

Friend Sub UpdateParity(ByVal c As String)
    Dim charBytes() As Byte
    If m_encKanji.InSubset(c) Then
        charBytes = m_shiftJISEncoding.GetBytes(c)
    Else
        charBytes = m_byteModeEncoding.GetBytes(c)
    End If

    Dim i As Long

    For i = 0 To UBound(charBytes)
        m_parity = m_parity Xor charBytes(i)
    Next
End Sub

Private Function SelectInitialMode(ByRef s As String, ByVal startIndex As Long) As EncodingMode
    If m_encKanji.InSubset(Mid$(s, startIndex, 1)) Then
        SelectInitialMode = EncodingMode.KANJI
        Exit Function
    End If

    If m_encByte.InExclusiveSubset(Mid$(s, startIndex, 1)) Then
        SelectInitialMode = EncodingMode.EIGHT_BIT_BYTE
        Exit Function
    End If

    Dim cnt As Long
    Dim flg As Boolean
    Dim i   As Long

    If m_encAlpha.InExclusiveSubset(Mid$(s, startIndex, 1)) Then
        cnt = 0
        flg = False

        For i = startIndex To Len(s)
            If m_encAlpha.InExclusiveSubset(Mid$(s, i, 1)) Then
                cnt = cnt + 1
            Else
                Exit For
            End If
        Next

        Select Case m_currSymbol.Version
            Case 1 To 9
                flg = cnt < 6
            Case 10 To 26
                flg = cnt < 7
            Case 27 To 40
                flg = cnt < 8
            Case Else
                Call Err.Raise(51)
        End Select

        If flg Then
            If (startIndex + cnt) <= Len(s) Then
                If m_encByte.InExclusiveSubset(Mid$(s, startIndex + cnt, 1)) Then
                    SelectInitialMode = EncodingMode.EIGHT_BIT_BYTE
                    Exit Function
                Else
                    SelectInitialMode = EncodingMode.ALPHA_NUMERIC
                    Exit Function
                End If
            Else
                SelectInitialMode = EncodingMode.ALPHA_NUMERIC
                Exit Function
            End If
        Else
            SelectInitialMode = EncodingMode.ALPHA_NUMERIC
            Exit Function
        End If
    End If

    Dim flg1 As Boolean
    Dim flg2 As Boolean

    If m_encNum.InSubset(Mid$(s, startIndex, 1)) Then
        cnt = 0
        flg1 = False
        flg2 = False

        For i = startIndex To Len(s)
            If m_encNum.InSubset(Mid$(s, i, 1)) Then
                cnt = cnt + 1
            Else
                Exit For
            End If
        Next

        Select Case m_currSymbol.Version
            Case 1 To 9
                flg1 = cnt < 4
                flg2 = cnt < 7
            Case 10 To 26
                flg1 = cnt < 4
                flg2 = cnt < 8
            Case 27 To 40
                flg1 = cnt < 5
                flg2 = cnt < 9
            Case Else
                Call Err.Raise(51)
        End Select

        If flg1 Then
            If (startIndex + cnt) <= Len(s) Then
                flg1 = m_encByte.InExclusiveSubset(Mid$(s, startIndex + cnt, 1))
            Else
                flg1 = False
            End If
        End If

        If flg2 Then
            If (startIndex + cnt) <= Len(s) Then
                flg2 = m_encAlpha.InExclusiveSubset(Mid$(s, startIndex + cnt, 1))
            Else
                flg2 = False
            End If
        End If

        If flg1 Then
            SelectInitialMode = EncodingMode.EIGHT_BIT_BYTE
            Exit Function
        ElseIf flg2 Then
            SelectInitialMode = EncodingMode.ALPHA_NUMERIC
            Exit Function
        Else
            SelectInitialMode = EncodingMode.NUMERIC
            Exit Function
        End If
    End If

    Call Err.Raise(51)
End Function

Private Function SelectModeWhileInNumericMode( _
    ByRef s As String, ByVal startIndex As Long) As EncodingMode

    If m_encKanji.InSubset(Mid$(s, startIndex, 1)) Then
        SelectModeWhileInNumericMode = EncodingMode.KANJI
        Exit Function
    End If
    
    If m_encByte.InExclusiveSubset(Mid$(s, startIndex, 1)) Then
        SelectModeWhileInNumericMode = EncodingMode.EIGHT_BIT_BYTE
        Exit Function
    End If
    
    If m_encAlpha.InExclusiveSubset(Mid$(s, startIndex, 1)) Then
        SelectModeWhileInNumericMode = EncodingMode.ALPHA_NUMERIC
        Exit Function
    End If

    SelectModeWhileInNumericMode = EncodingMode.NUMERIC
End Function

Private Function SelectModeWhileInAlphanumericMode( _
    ByRef s As String, ByVal startIndex As Long) As EncodingMode

    If m_encKanji.InSubset(Mid$(s, startIndex, 1)) Then
        SelectModeWhileInAlphanumericMode = EncodingMode.KANJI
        Exit Function
    ElseIf m_encByte.InExclusiveSubset(Mid$(s, startIndex, 1)) Then
        SelectModeWhileInAlphanumericMode = EncodingMode.EIGHT_BIT_BYTE
        Exit Function
    End If

    Dim cnt As Long
    cnt = 0
    
    Dim flg As Boolean
    flg = False
    
    Dim i As Long

    For i = startIndex To Len(s)
        If Not m_encAlpha.InSubset(Mid$(s, i, 1)) Then
            Exit For
        End If

        If m_encNum.InSubset(Mid$(s, i, 1)) Then
            cnt = cnt + 1
        Else
            flg = True
            Exit For
        End If
    Next

    If flg Then
        Select Case m_currSymbol.Version
            Case 1 To 9
                flg = cnt >= 13
            Case 10 To 26
                flg = cnt >= 15
            Case 27 To 40
                flg = cnt >= 17
            Case Else
                Call Err.Raise(51)
        End Select

        If flg Then
            SelectModeWhileInAlphanumericMode = EncodingMode.NUMERIC
            Exit Function
        End If
    End If

    SelectModeWhileInAlphanumericMode = EncodingMode.ALPHA_NUMERIC
End Function

Private Function SelectModeWhileInByteMode( _
    ByRef s As String, ByVal startIndex As Long) As EncodingMode

    If m_encKanji.InSubset(Mid$(s, startIndex, 1)) Then
        SelectModeWhileInByteMode = EncodingMode.KANJI
        Exit Function
    End If

    Dim cnt As Long
    Dim flg As Boolean
    Dim i   As Long

    For i = startIndex To Len(s)
        If Not m_encByte.InSubset(Mid$(s, i, 1)) Then
            Exit For
        End If

        If m_encNum.InSubset(Mid$(s, i, 1)) Then
            cnt = cnt + 1
        ElseIf m_encByte.InExclusiveSubset(Mid$(s, i, 1)) Then
            flg = True
            Exit For
        Else
            Exit For
        End If
    Next

    If flg Then
        Select Case m_currSymbol.Version
            Case 1 To 9
                flg = cnt >= 6
            Case 10 To 26
                flg = cnt >= 8
            Case 27 To 40
                flg = cnt >= 9
            Case Else
                Call Err.Raise(51)
        End Select

        If flg Then
            SelectModeWhileInByteMode = EncodingMode.NUMERIC
            Exit Function
        End If
    End If

    cnt = 0
    flg = False

    For i = startIndex To Len(s)
        If Not m_encByte.InSubset(Mid$(s, i, 1)) Then
            Exit For
        End If

        If m_encAlpha.InExclusiveSubset(Mid$(s, i, 1)) Then
            cnt = cnt + 1
        ElseIf m_encByte.InExclusiveSubset(Mid$(s, i, 1)) Then
            flg = True
            Exit For
        Else
            Exit For
        End If

        i = i + 1
    Next

    If flg Then
        Select Case m_currSymbol.Version
            Case 1 To 9
                flg = cnt >= 11
            Case 10 To 26
                flg = cnt >= 15
            Case 27 To 40
                flg = cnt >= 16
            Case Else
                Call Err.Raise(51)
        End Select

        If flg Then
            SelectModeWhileInByteMode = EncodingMode.ALPHA_NUMERIC
            Exit Function
        End If
    End If

    SelectModeWhileInByteMode = EncodingMode.EIGHT_BIT_BYTE
End Function
