Attribute VB_Name = "Constructors"
Option Private Module
Option Explicit

Public Function NewSymbols(ByVal ecLevel As ErrorCorrectionLevel, _
                ByVal maxVer As Long, _
                ByVal allowStructuredAppend As Boolean, _
                ByVal charsetName As String) As Symbols
    Dim ret As New Symbols
    Call ret.Init(ecLevel, maxVer, allowStructuredAppend, charsetName)
    Set NewSymbols = ret
End Function

Public Function NewSymbol(ByVal parentObj As Symbols) As Symbol
    Dim ret As New Symbol
    Call ret.Init(parentObj)
    Set NewSymbol = ret
End Function

Public Function NewKanjiEncoder(ByVal charsetName As String) As KanjiEncoder
    Dim ret As New KanjiEncoder
    Call ret.Init(charsetName)
    Set NewKanjiEncoder = ret
End Function

Public Function NewByteEncoder(ByVal charsetName As String) As ByteEncoder
    Dim ret As New ByteEncoder
    Call ret.Init(charsetName)
    Set NewByteEncoder = ret
End Function

Public Function NewTextStream(ByVal charsetName As String) As QRCodeLib.TextStream
    Dim ret As New QRCodeLib.TextStream
    Call ret.Init(charsetName)
    Set NewTextStream = ret
End Function

Public Function NewPoint(ByVal x As Long, ByVal y As Long) As Point
    Dim ret As New Point
    Call ret.Init(x, y)
    Set NewPoint = ret
End Function
