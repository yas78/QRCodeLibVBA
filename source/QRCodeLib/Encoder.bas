Attribute VB_Name = "Encoder"
Option Private Module
Option Explicit

Public Function Create(ByVal encMode As EncodingMode, _
                       ByVal byteModeEncoding As String) As IQRCodeEncoder
    Dim ret As IQRCodeEncoder
    Dim enc As ByteEncoder

    Select Case encMode
        Case EncodingMode.NUMERIC
            Set ret = New NumericEncoder
        Case EncodingMode.ALPHA_NUMERIC
            Set ret = New AlphanumericEncoder
        Case EncodingMode.EIGHT_BIT_BYTE
            Set ret = New ByteEncoder
            Set enc = ret
            Call enc.Init(byteModeEncoding)
        Case EncodingMode.KANJI
            Set ret = New KanjiEncoder
        Case Else
            Call Err.Raise(5)
    End Select

    Set Create = ret
End Function
