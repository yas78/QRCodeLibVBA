Attribute VB_Name = "Factory"
Option Explicit

Public Function NewSymbols(Optional ByVal maxVer As Long = Constants.MAX_VERSION, _
                           Optional ByVal ecLevel As ErrorCorrectionLevel = ErrorCorrectionLevel.M, _
                           Optional ByVal allowStructuredAppend As Boolean = False) As Symbols
    
    Dim sbls As New Symbols
    
    Call sbls.Initialize(maxVer, ecLevel, allowStructuredAppend)
    Set NewSymbols = sbls
    
End Function
