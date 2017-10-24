Attribute VB_Name = "VersionInfo"
'----------------------------------------------------------------------------------------
' 型番情報
'----------------------------------------------------------------------------------------
Option Private Module
Option Explicit

Private m_versionInfoValues() As Variant

Private m_initialized As Boolean

'----------------------------------------------------------------------------------------
' (概要)
'  型番情報を配置します｡
'----------------------------------------------------------------------------------------
Public Sub Place(ByRef moduleMatrix() As Variant, ByVal ver As Long)

#If [DEBUG] Then
    Debug.Assert ver >= 7 And _
                 ver <= Constants.MAX_VERSION
#End If

    Call Initialize

    Dim numModulesPerSide As Long
    numModulesPerSide = UBound(moduleMatrix) + 1
    
    Dim versionInfoValue As Long
    versionInfoValue = m_versionInfoValues(ver)
    
    Dim p1  As Long
    p1 = 0
    
    Dim p2  As Long
    p2 = numModulesPerSide - 11
    
    Dim i   As Long
    Dim v   As Long
    
    For i = 0 To 17
        v = IIf((versionInfoValue And 2 ^ i) > 0, 3, -3)
        
        moduleMatrix(p1)(p2) = v
        moduleMatrix(p2)(p1) = v
        
        p2 = p2 + 1

        If i Mod 3 = 2 Then
            p1 = p1 + 1
            p2 = numModulesPerSide - 11
        End If
        
    Next

End Sub
 
'----------------------------------------------------------------------------------------
' (概要)
'  型番情報の予約領域を配置します｡
'----------------------------------------------------------------------------------------
Public Sub PlaceTempBlank(ByRef moduleMatrix() As Variant)

    Dim numModulesPerSide As Long
    numModulesPerSide = UBound(moduleMatrix) + 1
    
    Dim i As Long
    Dim j As Long

    For i = 0 To 5
        For j = numModulesPerSide - 11 To numModulesPerSide - 9
            moduleMatrix(i)(j) = -3  ' 右上
            moduleMatrix(j)(i) = -3  ' 左下
        Next
    Next
    
End Sub

'----------------------------------------------------------------------------------------
' (概要)
'  オブジェクトを初期化します。
'----------------------------------------------------------------------------------------
Private Sub Initialize()

    If m_initialized Then Exit Sub

    m_initialized = True

    ' 型番情報
    m_versionInfoValues = Array( _
        -1, -1, -1, -1, -1, -1, -1, _
        &H7C94&, &H85BC&, &H9A99&, &HA4D3&, &HBBF6&, &HC762&, &HD847&, &HE60D&, _
        &HF928&, &H10B78, &H1145D, &H12A17, &H13532, &H149A6, &H15683, &H168C9, _
        &H177EC, &H18EC4, &H191E1, &H1AFAB, &H1B08E, &H1CC1A, &H1D33F, &H1ED75, _
        &H1F250, &H209D5, &H216F0, &H228BA, &H2379F, &H24B0B, &H2542E, &H26A64, _
        &H27541, &H28C69 _
    )

End Sub
