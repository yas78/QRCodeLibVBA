VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VersionInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const VAL As Long = Values.VERSION_INFO

Private m_values() As Long

Private Sub Class_Initialize()
    m_values = LongArray( _
        -1, -1, -1, -1, -1, -1, -1, _
        &H7C94&, &H85BC&, &H9A99&, &HA4D3&, &HBBF6&, &HC762&, &HD847&, &HE60D&, _
        &HF928&, &H10B78, &H1145D, &H12A17, &H13532, &H149A6, &H15683, &H168C9, _
        &H177EC, &H18EC4, &H191E1, &H1AFAB, &H1B08E, &H1CC1A, &H1D33F, &H1ED75, _
        &H1F250, &H209D5, &H216F0, &H228BA, &H2379F, &H24B0B, &H2542E, &H26A64, _
        &H27541, &H28C69 _
    )
End Sub

Public Sub Place(ByVal ver As Long, ByRef moduleMatrix() As Variant)
    Dim numModulesPerSide As Long
    numModulesPerSide = UBound(moduleMatrix) + 1

    Dim versionInfoValue As Long
    versionInfoValue = m_values(ver)

    Dim p1 As Long
    p1 = 0

    Dim p2 As Long
    p2 = numModulesPerSide - 11

    Dim i As Long
    Dim v As Long
    For i = 0 To 17
        v = IIf((versionInfoValue And 2 ^ i) > 0, VAL, -VAL)

        moduleMatrix(p1)(p2) = v
        moduleMatrix(p2)(p1) = v

        p2 = p2 + 1

        If i Mod 3 = 2 Then
            p1 = p1 + 1
            p2 = numModulesPerSide - 11
        End If
    Next
End Sub

Public Sub PlaceTempBlank(ByRef moduleMatrix() As Variant)
    Dim numModulesPerSide As Long
    numModulesPerSide = UBound(moduleMatrix) + 1

    Dim i As Long
    Dim j As Long
    For i = 0 To 5
        For j = numModulesPerSide - 11 To numModulesPerSide - 9
            moduleMatrix(i)(j) = -VAL
            moduleMatrix(j)(i) = -VAL
        Next
    Next
End Sub
