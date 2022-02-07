Attribute VB_Name = "Factory"
Option Explicit

Public Function CreateSymbols( _
    Optional ByVal ecLevel As ErrorCorrectionLevel = ErrorCorrectionLevel.M, _
    Optional ByVal maxVer As Long = Constants.MAX_VERSION, _
    Optional ByVal allowStructuredAppend As Boolean = False, _
    Optional ByVal charsetName As String = Charset.SHIFT_JIS, _
    Optional ByVal fixedSize As Boolean = False) As Symbols

    If Not (ErrorCorrectionLevel.L <= ecLevel And ecLevel <= ErrorCorrectionLevel.H) Then Call Err.Raise(5)
    If Not (Constants.MIN_VERSION <= maxVer And maxVer <= Constants.MAX_VERSION) Then Call Err.Raise(5)

    Dim charEncoding As New Encoding
    Call charEncoding.Init(charsetName)

    Dim minVer As Long
    minVer = IIf(fixedSize, maxVer, Constants.MIN_VERSION)

    Dim sbls As New Symbols
    Call sbls.Init(ecLevel, minVer, maxVer, allowStructuredAppend, charEncoding)

    Set CreateSymbols = sbls
End Function
