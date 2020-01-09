Attribute VB_Name = "CharCountIndicator"
Option Private Module
Option Explicit

Public Function GetLength( _
    ByVal ver As Long, ByVal encMode As EncodingMode) As Long

    Select Case ver
        Case 1 To 9
            Select Case encMode
                Case EncodingMode.NUMERIC
                    GetLength = 10
                Case EncodingMode.ALPHA_NUMERIC
                    GetLength = 9
                Case EncodingMode.EIGHT_BIT_BYTE
                    GetLength = 8
                Case EncodingMode.KANJI
                    GetLength = 8
                Case Else
                    Call Err.Raise(5)
            End Select

        Case 10 To 26
            Select Case encMode
                Case EncodingMode.NUMERIC
                    GetLength = 12
                Case EncodingMode.ALPHA_NUMERIC
                    GetLength = 11
                Case EncodingMode.EIGHT_BIT_BYTE
                    GetLength = 16
                Case EncodingMode.KANJI
                    GetLength = 10
                Case Else
                    Call Err.Raise(5)
            End Select

        Case 27 To 40
            Select Case encMode
                Case EncodingMode.NUMERIC
                    GetLength = 14
                Case EncodingMode.ALPHA_NUMERIC
                    GetLength = 13
                Case EncodingMode.EIGHT_BIT_BYTE
                    GetLength = 16
                Case EncodingMode.KANJI
                    GetLength = 12
                Case Else
                    Call Err.Raise(5)
            End Select
    Case Else
        Call Err.Raise(5)
    End Select
End Function
