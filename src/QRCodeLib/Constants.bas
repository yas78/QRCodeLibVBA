Attribute VB_Name = "Constants"
Option Private Module
Option Explicit

Public Const MIN_VERSION As Long = 1
Public Const MAX_VERSION As Long = 40

Public Enum EncodingMode
    UNKNOWN
    NUMERIC
    ALPHA_NUMERIC
    EIGHT_BIT_BYTE
    KANJI
End Enum
