Attribute VB_Name = "Charset"
Option Private Module
Option Explicit

Public Const SHIFT_JIS As String = "Shift_JIS"

Public Function IsJP(ByVal charsetName As String) As Boolean
    IsJP = LCase$(charsetName) = LCase$(SHIFT_JIS)
End Function
