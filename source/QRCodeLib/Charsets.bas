Attribute VB_Name = "Charsets"
Option Private Module
Option Explicit

Public Const SHIFT_JIS As String = "shift_jis"
Public Const GB2312 As String = "gb2312"
Public Const EUC_KR As String = "euc-kr"

Public Function IsCJK(ByVal charsetName As String) As Boolean
    Dim v As Variant
    
    For Each v In CJKCharsetNames()
        If LCase(charsetName) = LCase(v) Then
            IsCJK = True
            Exit Function
        End If
    Next
    
    IsCJK = False
End Function

Private Function CJKCharsetNames() As Variant()
    Dim ret() As Variant
    ret = Array(SHIFT_JIS, GB2312, EUC_KR)
    CJKCharsetNames = ret
End Function
