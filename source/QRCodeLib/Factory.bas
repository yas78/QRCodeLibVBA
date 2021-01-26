Attribute VB_Name = "Factory"
Option Explicit

Public Function CreateSymbols( _
    Optional ByVal ecLevel As ErrorCorrectionLevel = ErrorCorrectionLevel.M, _
    Optional ByVal maxVer As Long = MAX_VERSION, _
    Optional ByVal allowStructuredAppend As Boolean = False, _
    Optional ByVal charsetName As String = "shift_jis") As Symbols

    Select Case ecLevel
        Case ErrorCorrectionLevel.L To ErrorCorrectionLevel.H
            ' NOP
        Case Else
            Call Err.Raise(5)
    End Select

    If Not (1 <= maxVer And maxVer <= 40) Then Call Err.Raise(5)

    Dim sbls As Symbols
    Set sbls = NewSymbols(ecLevel, maxVer, allowStructuredAppend, charsetName)
    
    Set CreateSymbols = sbls
End Function
