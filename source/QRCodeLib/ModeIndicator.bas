Attribute VB_Name = "ModeIndicator"
'------------------------------------------------------------------------------
' モード指示子
'------------------------------------------------------------------------------
Option Private Module
Option Explicit


Public Const Length As Long = 4

Public Const TERMINATOR_VALUE           As Long = &H0
Public Const NUMERIC_VALUE              As Long = &H1
Public Const ALPAHNUMERIC_VALUE         As Long = &H2
Public Const STRUCTURED_APPEND_VALUE    As Long = &H3
Public Const BYTE_VALUE                 As Long = &H4
Public Const KANJI_VALUE                As Long = &H8
