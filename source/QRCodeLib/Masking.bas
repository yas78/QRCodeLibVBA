Attribute VB_Name = "Masking"
'------------------------------------------------------------------------------
' マスク
'------------------------------------------------------------------------------
Option Private Module
Option Explicit


'------------------------------------------------------------------------------
' (概要)
'  マスクを適用します。
'
' (パラメータ)
'  moduleMatrix : シンボルの明暗パターン
'  ver          : 型番
'  ecLevel      : 誤り訂正レベル
'
' (戻り値)
'  適用されたマスクパターン参照子
'------------------------------------------------------------------------------
Public Function Apply(ByVal ver As Long, _
                      ByVal ecLevel As ErrorCorrectionLevel, _
                      ByRef moduleMatrix() As Variant) As Long
    Dim minPenalty As Long
    minPenalty = &H7FFFFFFF

    Dim temp()  As Variant
    Dim penalty As Long
    Dim maskPatternReference As Long
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
            maskPatternReference = i
            maskedMatrix = temp
        End If
    Next

    moduleMatrix = maskedMatrix
    Apply = maskPatternReference
End Function

'------------------------------------------------------------------------------
' (概要)
'  マスクパターンを適用したシンボルデータを返します。
'
' (パラメータ)
'  moduleMatrix()       : シンボルの明暗パターン
'  maskPatternReference : マスクパターン参照子を表す0から7までの値
'------------------------------------------------------------------------------
Private Sub Mask(ByVal maskPatternReference As Long, ByRef moduleMatrix() As Variant)
    Dim condition As IMaskingCondition
    Set condition = GetCondition(maskPatternReference)

    Dim r As Long
    Dim c As Long

    For r = 0 To UBound(moduleMatrix)
        For c = 0 To UBound(moduleMatrix(r))
            If Math.Abs(moduleMatrix(r)(c)) = 1 Then
                If condition.Evaluate(r, c) Then
                    moduleMatrix(r)(c) = moduleMatrix(r)(c) * -1
                End If
            End If
        Next
    Next
End Sub

'------------------------------------------------------------------------------
' (概要)
'  マスク条件を返します。
'------------------------------------------------------------------------------
Private Function GetCondition(ByVal maskPatternReference As Long) As IMaskingCondition
    Dim ret As IMaskingCondition

    Select Case maskPatternReference
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
