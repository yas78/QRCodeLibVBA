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
Public Function Apply(ByRef moduleMatrix() As Variant, _
                      ByVal ver As Long, _
                      ByVal ecLevel As ErrorCorrectionLevel) As Long

    Dim maskPatternReference As Long
    maskPatternReference = SelectMaskPattern(moduleMatrix, ver, ecLevel)
    Call Mask(moduleMatrix, maskPatternReference)

    Apply = maskPatternReference

End Function

'------------------------------------------------------------------------------
' (概要)
'  マスクパターンを決定します。
'
' (パラメータ)
'  moduleMatrix : シンボルの明暗パターン
'  ver          : 型番
'  ecLevel      : 誤り訂正レベル
'
' (戻り値)
'  マスクパターン参照子
'------------------------------------------------------------------------------
Private Function SelectMaskPattern(ByRef moduleMatrix() As Variant, _
                                   ByVal ver As Long, _
                                   ByVal ecLevel As ErrorCorrectionLevel) As Long

    Dim minPenalty As Long
    minPenalty = &H7FFFFFFF

    Dim ret As Long
    ret = 0

    Dim temp()  As Variant
    Dim penalty As Long
    Dim maskPatternReference As Long

    For maskPatternReference = 0 To 7
        temp = moduleMatrix

        Call Mask(temp, maskPatternReference)

        Call FormatInfo.Place(temp, ecLevel, maskPatternReference)

        If ver >= 7 Then
            Call VersionInfo.Place(temp, ver)
        End If

        penalty = MaskingPenaltyScore.CalcTotal(temp)

        If penalty < minPenalty Then
            minPenalty = penalty
            ret = maskPatternReference
        End If
    Next

    SelectMaskPattern = ret

End Function


'------------------------------------------------------------------------------
' (概要)
'  マスクパターンを適用したシンボルデータを返します。
'
' (パラメータ)
'  moduleMatrix()       : シンボルの明暗パターン
'  maskPatternReference : マスクパターン参照子を表す0から7までの値
'------------------------------------------------------------------------------
Private Sub Mask(ByRef moduleMatrix() As Variant, ByVal maskPatternReference As Long)

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
            Set ret = New Masking0Condition

        Case 1
            Set ret = New Masking1Condition

        Case 2
            Set ret = New Masking2Condition

        Case 3
            Set ret = New Masking3Condition

        Case 4
            Set ret = New Masking4Condition

        Case 5
            Set ret = New Masking5Condition

        Case 6
            Set ret = New Masking6Condition

        Case 7
            Set ret = New Masking7Condition

        Case Else
            Call Err.Raise(5)

    End Select

    Set GetCondition = ret

End Function
