Attribute VB_Name = "Encoder"
Option Private Module
Option Explicit

Public Function Create(ByVal encMode As EncodingMode, _
                       ByVal charsetName As String) As IQRCodeEncoder
    Dim ret As IQRCodeEncoder

    Select Case encMode
        Case EncodingMode.NUMERIC
            Set ret = New NumericEncoder
        Case EncodingMode.ALPHA_NUMERIC
            Set ret = New AlphanumericEncoder
        Case EncodingMode.EIGHT_BIT_BYTE
            If Len(charsetName) = 0 Then Call Err.Raise(5)
            Set ret = NewByteEncoder(charsetName)
        Case EncodingMode.KANJI
            If Encoding.IsCJK(charsetName) Then
                Set ret = NewKanjiEncoder(charsetName)
            Else
                Call Err.Raise(5)
            End If
        Case Else
            Call Err.Raise(5)
    End Select

    Set Create = ret
End Function
