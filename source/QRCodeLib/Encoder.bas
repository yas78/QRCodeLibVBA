Attribute VB_Name = "Encoder"
Option Private Module
Option Explicit

Public Function Create(ByVal encMode As EncodingMode, _
                       ByVal charEncoding As Encoding) As IQRCodeEncoder
    If charEncoding Is Nothing Then Call Err.Raise(5)
    
    Dim ret As IQRCodeEncoder

    Select Case encMode
        Case EncodingMode.NUMERIC
            Set ret = New NumericEncoder
        Case EncodingMode.ALPHA_NUMERIC
            Set ret = New AlphanumericEncoder
        Case EncodingMode.EIGHT_BIT_BYTE
            Set ret = New ByteEncoder
        Case EncodingMode.KANJI
            If Charsets.IsCJK(charEncoding.Charset) Then
                Set ret = New KanjiEncoder
            Else
                Call Err.Raise(5)
            End If
        Case Else
            Call Err.Raise(5)
    End Select
    
    Call ret.Init(charEncoding)
    Set Create = ret
End Function
