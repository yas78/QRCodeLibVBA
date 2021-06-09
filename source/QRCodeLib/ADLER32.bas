Attribute VB_Name = "ADLER32"
Option Private Module
Option Explicit

Public Function Checksum(ByRef data() As Byte) As Long
    Checksum = Update(1, data)
End Function

Public Function Update(ByVal adler As Long, ByRef data() As Byte) As Long
    Dim s1 As Long
    s1 = adler And &HFFFF&

    Dim s2  As Long
    s2 = (adler \ 2 ^ 16) And &HFFFF&

    Dim n As Long
    For n = 0 To UBound(data)
        s1 = (s1 + data(n)) Mod 65521
        s2 = (s2 + s1) Mod 65521
    Next

    Dim temp As Long

    If (s2 And &H8000&) > 0 Then
        temp = ((s2 And &H7FFF&) * 2 ^ 16) Or &H80000000
    Else
        temp = s2 * 2 ^ 16
    End If

    Update = temp + s1
End Function
