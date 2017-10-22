Attribute VB_Name = "Factory"
Option Explicit

'----------------------------------------------------------------------------------------
' (概要)
'  Symbolsクラスのインスタンスを生成します。
'
' (パラメータ)
'　maxVer                : 型番の上限
'  ecLevel               : 誤り訂正レベル
'  allowStructuredAppend : 複数シンボルへの分割を許可するには True を指定します。
'  byteModeCharsetName   : バイトモードの文字コードを"Shift_JIS" または "UTF-8" で指定します。
'----------------------------------------------------------------------------------------
Public Function NewSymbols(Optional ByVal maxVer As Long = Constants.MAX_VERSION, _
                           Optional ByVal ecLevel As ErrorCorrectionLevel = ErrorCorrectionLevel.M, _
                           Optional ByVal allowStructuredAppend As Boolean = False, _
                           Optional ByVal byteModeCharsetName As String = "Shift_JIS") As Symbols
    
    If LCase(byteModeCharsetName) <> "shift_jis" And _
       LCase(byteModeCharsetName) <> "utf-8" Then
        Err.Raise 5
    End If
    
    Dim sbls As New Symbols
    
    Call sbls.Initialize(maxVer, ecLevel, allowStructuredAppend, byteModeCharsetName)
    Set NewSymbols = sbls
    
End Function
