Attribute VB_Name = "Crc32"
Option Private Module
Option Explicit

Private m_crcTable(255)    As Long
Private m_crcTableComputed As Boolean

Public Function Checksum(ByRef data() As Byte) As Long
    Checksum = Update(0, data)
End Function

Public Function Update(ByVal crc As Long, ByRef data() As Byte) As Long
    Dim c As Long
    c = crc Xor &HFFFFFFFF

    If Not m_crcTableComputed Then
        Call MakeCrcTable
    End If

    Dim n As Long
    For n = 0 To UBound(data)
        c = m_crcTable((c Xor data(n)) And &HFF) Xor _
                (((c And &HFFFFFF00) \ 2 ^ 8) And &HFFFFFF)
    Next

    Update = c Xor &HFFFFFFFF
End Function

Private Sub MakeCrcTable()
    Dim c As Long

    Dim k As Long
    Dim n As Long
    For n = 0 To 255
        c = n
        For k = 0 To 7
            If c And 1 Then
                c = &HEDB88320 Xor ((c And &HFFFFFFFE) \ 2 And &H7FFFFFFF)
            Else
                c = (c \ 2) And &H7FFFFFFF
            End If
        Next

        m_crcTable(n) = c
    Next

    m_crcTableComputed = True
End Sub
