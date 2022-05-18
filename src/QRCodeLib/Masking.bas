Attribute VB_Name = "Masking"
Option Private Module
Option Explicit

Public Function Apply(ByVal ver As Long, _
                      ByVal ecLevel As ErrorCorrectionLevel, _
                      ByRef moduleMatrix() As Variant) As Long
    Dim minPenalty As Long
    minPenalty = &H7FFFFFFF

    Dim temp() As Variant
    Dim penalty As Long
    Dim maskPattern As Long
    Dim maskedMatrix() As Variant

    Dim i As Long
    For i = 0 To 7
        temp = moduleMatrix

        Call Mask(i, temp)
        Call FormatInfo.Place(ecLevel, i, temp)

        If ver >= 7 Then
            Call VersionInfo.Place(ver, temp)
        End If

        penalty = MaskingPenaltyScore.CalcTotal(temp)

        If penalty < minPenalty Then
            minPenalty = penalty
            maskPattern = i
            maskedMatrix = temp
        End If
    Next

    moduleMatrix = maskedMatrix
    Apply = maskPattern
End Function

Private Sub Mask(ByVal maskPattern As Long, ByRef moduleMatrix() As Variant)
    Dim condition As IMaskingCondition
    Set condition = GetCondition(maskPattern)

    Dim r As Long
    Dim c As Long
    For r = 0 To UBound(moduleMatrix)
        For c = 0 To UBound(moduleMatrix(r))
            If Math.Abs(moduleMatrix(r)(c)) = Values.WORD Then
                If condition.Evaluate(r, c) Then
                    moduleMatrix(r)(c) = -(moduleMatrix(r)(c))
                End If
            End If
        Next
    Next
End Sub

Private Function GetCondition(ByVal maskPattern As Long) As IMaskingCondition
    Dim ret As IMaskingCondition

    Select Case maskPattern
        Case 0
            Set ret = New MaskingCondition0
        Case 1
            Set ret = New MaskingCondition1
        Case 2
            Set ret = New MaskingCondition2
        Case 3
            Set ret = New MaskingCondition3
        Case 4
            Set ret = New MaskingCondition4
        Case 5
            Set ret = New MaskingCondition5
        Case 6
            Set ret = New MaskingCondition6
        Case 7
            Set ret = New MaskingCondition7
        Case Else
            Call Err.Raise(5)
    End Select

    Set GetCondition = ret
End Function
