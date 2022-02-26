VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RSBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_totalNumbers() As Variant

Private Sub Class_Initialize()
    Dim ecLevelL() As Long
    ecLevelL = LongArray( _
         0, _
         1, 1, 1, 1, 1, 2, 2, 2, 2, 4, _
         4, 4, 4, 4, 6, 6, 6, 6, 7, 8, _
         8, 9, 9, 10, 12, 12, 12, 13, 14, 15, _
        16, 17, 18, 19, 19, 20, 21, 22, 24, 25 _
    )

    Dim ecLevelM() As Long
    ecLevelM = LongArray( _
         0, _
         1, 1, 1, 2, 2, 4, 4, 4, 5, 5, _
         5, 8, 9, 9, 10, 10, 11, 13, 14, 16, _
        17, 17, 18, 20, 21, 23, 25, 26, 28, 29, _
        31, 33, 35, 37, 38, 40, 43, 45, 47, 49 _
    )

    Dim ecLevelQ() As Long
    ecLevelQ = LongArray( _
         0, _
         1, 1, 2, 2, 4, 4, 6, 6, 8, 8, _
         8, 10, 12, 16, 12, 17, 16, 18, 21, 20, _
        23, 23, 25, 27, 29, 34, 34, 35, 38, 40, _
        43, 45, 48, 51, 53, 56, 59, 62, 65, 68 _
    )

    Dim ecLevelH() As Long
    ecLevelH = LongArray( _
         0, _
         1, 1, 2, 4, 4, 4, 5, 6, 8, 8, _
        11, 11, 16, 16, 18, 16, 19, 21, 25, 25, _
        25, 34, 30, 32, 35, 37, 40, 42, 45, 48, _
        51, 54, 57, 60, 63, 66, 70, 74, 77, 81 _
    )

    m_totalNumbers = Array(ecLevelL, ecLevelM, ecLevelQ, ecLevelH)
End Sub

Public Function GetTotalNumber(ByVal ecLevel As ErrorCorrectionLevel, _
                               ByVal ver As Long, _
                               ByVal preceding As Boolean) As Long
    Dim dataWordCapacity As Long
    dataWordCapacity = DataCodeword.GetTotalNumber(ecLevel, ver)

    Dim blockCount() As Long
    blockCount = m_totalNumbers(ecLevel)

    If preceding Then
        GetTotalNumber = blockCount(ver) - (dataWordCapacity Mod blockCount(ver))
    Else
        GetTotalNumber = dataWordCapacity Mod blockCount(ver)
    End If
End Function

Public Function GetNumberDataCodewords(ByVal ecLevel As ErrorCorrectionLevel, _
                                       ByVal ver As Long, _
                                       ByVal preceding As Boolean) As Long
    Dim numDataCodewords As Long
    numDataCodewords = DataCodeword.GetTotalNumber(ecLevel, ver)

    Dim numBlocks As Long
    numBlocks = m_totalNumbers(ecLevel)(ver)

    Dim numPreBlockCodewords As Long
    numPreBlockCodewords = numDataCodewords \ numBlocks

    Dim numPreBlocks As Long
    Dim numFolBlocks As Long

    Dim ret As Long

    If preceding Then
        ret = numPreBlockCodewords
    Else
        numPreBlocks = GetTotalNumber(ecLevel, ver, True)
        numFolBlocks = GetTotalNumber(ecLevel, ver, False)

        If numFolBlocks > 0 Then
            ret = (numDataCodewords - numPreBlockCodewords * numPreBlocks) \ numFolBlocks
        Else
            ret = 0
        End If
    End If

    GetNumberDataCodewords = ret
End Function

Public Function GetNumberECCodewords(ByVal ecLevel As ErrorCorrectionLevel, _
                                     ByVal ver As Long) As Long
    Dim numDataCodewords As Long
    numDataCodewords = DataCodeword.GetTotalNumber(ecLevel, ver)

    Dim numBlocks As Long
    numBlocks = m_totalNumbers(ecLevel)(ver)

    GetNumberECCodewords = _
        (Codeword.GetTotalNumber(ver) \ numBlocks) - _
            (numDataCodewords \ numBlocks)
End Function
