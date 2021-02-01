Attribute VB_Name = "Factory"
Option Explicit

Public Function CreateSymbols( _
    Optional ByVal ecLevel As ErrorCorrectionLevel = ErrorCorrectionLevel.M, _
    Optional ByVal maxVer As Long = MAX_VERSION, _
    Optional ByVal allowStructuredAppend As Boolean = False, _
    Optional ByVal charsetName As String = Charsets.SHIFT_JIS) As Symbols

    Select Case ecLevel
        Case ErrorCorrectionLevel.L To ErrorCorrectionLevel.H
            ' NOP
        Case Else
            Call Err.Raise(5)
    End Select

    If Not (1 <= maxVer And maxVer <= 40) Then Call Err.Raise(5)

    Dim charEncoding As New Encoding
    Call charEncoding.Init(charsetName)

    Dim sbls As New Symbols
    Call sbls.Init(ecLevel, maxVer, allowStructuredAppend, charEncoding)

    Set CreateSymbols = sbls
End Function
