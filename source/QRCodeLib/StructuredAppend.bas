Attribute VB_Name = "StructuredAppend"
'------------------------------------------------------------------------------
' 構造的連接
'------------------------------------------------------------------------------
Option Private Module
Option Explicit

' パリティデータのビット数
Public Const PARITY_DATA_LENGTH As Long = 8

' ヘッダーのビット数
Public Const HEADER_LENGTH As Long = _
    ModeIndicator.Length + _
    SymbolSequenceIndicator.POSITION_LENGTH + _
    SymbolSequenceIndicator.TOTAL_NUMBER_LENGTH + _
    PARITY_DATA_LENGTH
