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
'----------------------------------------------------------------------------------------
Public Function NewSymbols(Optional ByVal maxVer As Long = Constants.MAX_VERSION, _
                           Optional ByVal ecLevel As ErrorCorrectionLevel = ErrorCorrectionLevel.M, _
                           Optional ByVal allowStructuredAppend As Boolean = False) As Symbols
    
    Dim sbls As New Symbols
    
    Call sbls.Initialize(maxVer, ecLevel, allowStructuredAppend)
    Set NewSymbols = sbls
    
End Function
