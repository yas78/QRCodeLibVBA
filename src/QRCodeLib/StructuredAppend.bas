Attribute VB_Name = "StructuredAppend"
Option Private Module
Option Explicit

Public Const PARITY_DATA_LENGTH As Long = 8

Public Const HEADER_LENGTH As Long = _
    ModeIndicator.Length + _
    SymbolSequenceIndicator.POSITION_LENGTH + _
    SymbolSequenceIndicator.TOTAL_NUMBER_LENGTH + _
    PARITY_DATA_LENGTH
