Attribute VB_Name = "Factory"
Option Explicit

'------------------------------------------------------------------------------
' (概要)
'  Symbolsクラスのインスタンスを生成します。
'
' (パラメータ)
'  ecLevel               : 誤り訂正レベル
'　maxVer                : 型番の上限
'  allowStructuredAppend : 複数シンボルへの分割を許可するには True を指定します。
'  byteModeCharsetName   : バイトモードの文字コードを指定します。
'------------------------------------------------------------------------------
Public Function CreateSymbols( _
    Optional ByVal ecLevel As ErrorCorrectionLevel = ErrorCorrectionLevel.M, _
    Optional ByVal maxVer As Long = Constants.MAX_VERSION, _
    Optional ByVal allowStructuredAppend As Boolean = False, _
    Optional ByVal byteModeCharsetName As String = "Shift_JIS") As Symbols
    
    Dim sbls As New Symbols
    
    Call sbls.Initialize(ecLevel, maxVer, allowStructuredAppend, byteModeCharsetName)
    Set CreateSymbols = sbls
    
End Function
