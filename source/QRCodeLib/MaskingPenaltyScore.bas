Attribute VB_Name = "MaskingPenaltyScore"
'------------------------------------------------------------------------------
' マスクされたシンボルの失点評価
'------------------------------------------------------------------------------
Option Private Module
Option Explicit

'------------------------------------------------------------------------------
' (概要)
'  マスクパターン失点の合計を返します。
'------------------------------------------------------------------------------
Public Function CalcTotal(ByRef moduleMatrix() As Variant) As Long

    Dim total   As Long
    Dim penalty As Long

    penalty = CalcAdjacentModulesInSameColor(moduleMatrix)
    total = total + penalty

    penalty = CalcBlockOfModulesInSameColor(moduleMatrix)
    total = total + penalty

    penalty = CalcModuleRatio(moduleMatrix)
    total = total + penalty

    penalty = CalcProportionOfDarkModules(moduleMatrix)
    total = total + penalty

    CalcTotal = total

End Function


'------------------------------------------------------------------------------
' (概要)
'  行／列の同色隣接モジュールパターンの失点を計算します。
'------------------------------------------------------------------------------
Private Function CalcAdjacentModulesInSameColor(ByRef moduleMatrix() As Variant) As Long

    Dim penalty As Long
    penalty = 0

    penalty = penalty + CalcAdjacentModulesInRowInSameColor(moduleMatrix)
    penalty = penalty + CalcAdjacentModulesInRowInSameColor(ArrayUtil.Rotate90(moduleMatrix))

    CalcAdjacentModulesInSameColor = penalty

End Function

'------------------------------------------------------------------------------
' (概要)
'  行の同色隣接モジュールパターンの失点を計算します。
'------------------------------------------------------------------------------
Private Function CalcAdjacentModulesInRowInSameColor(ByRef moduleMatrix() As Variant) As Long

    Dim penalty As Long
    penalty = 0

    Dim rowArray As Variant
    Dim i As Long
    Dim cnt As Long

    For Each rowArray In moduleMatrix
        cnt = 1

        For i = 0 To UBound(rowArray) - 1
            If (rowArray(i) > 0) = (rowArray(i + 1) > 0) Then
                cnt = cnt + 1
            Else
                If cnt >= 5 Then
                    penalty = penalty + (3 + (cnt - 5))
                End If

                cnt = 1
            End If
        Next

        If cnt >= 5 Then
            penalty = penalty + (3 + (cnt - 5))
        End If
    Next

    CalcAdjacentModulesInRowInSameColor = penalty

End Function

'------------------------------------------------------------------------------
' (概要)
'  2x2の同色モジュールパターンの失点を計算します。
'------------------------------------------------------------------------------
Private Function CalcBlockOfModulesInSameColor(ByRef moduleMatrix() As Variant) As Long

    Dim penalty     As Long
    Dim isSameColor As Boolean
    Dim r           As Long
    Dim c           As Long
    Dim temp        As Boolean

    For r = 0 To UBound(moduleMatrix) - 1
        For c = 0 To UBound(moduleMatrix(r)) - 1
            temp = moduleMatrix(r)(c) > 0

            If (moduleMatrix(r + 0)(c + 1) > 0 = temp) And _
               (moduleMatrix(r + 1)(c + 0) > 0 = temp) And _
               (moduleMatrix(r + 1)(c + 1) > 0 = temp) Then
                penalty = penalty + 3
            End If

        Next
    Next

    CalcBlockOfModulesInSameColor = penalty

End Function

'------------------------------------------------------------------------------
' (概要)
'  行／列における1 : 1 : 3 : 1 : 1 比率パターンの失点を計算します。
'------------------------------------------------------------------------------
Private Function CalcModuleRatio(ByRef moduleMatrix() As Variant) As Long

    Dim moduleMatrixTemp() As Variant
    moduleMatrixTemp = QuietZone.Place(moduleMatrix)

    Dim penalty As Long
    penalty = 0

    penalty = penalty + CalcModuleRatioInRow(moduleMatrixTemp)
    penalty = penalty + CalcModuleRatioInRow(ArrayUtil.Rotate90(moduleMatrixTemp))

    CalcModuleRatio = penalty

End Function


'------------------------------------------------------------------------------
' (概要)
'  行の1 : 1 : 3 : 1 : 1 比率のパターンを評価し、失点を返します。
'------------------------------------------------------------------------------
Private Function CalcModuleRatioInRow(ByRef moduleMatrix() As Variant) As Long

    Dim penalty As Long

    Dim ratio3Ranges As Collection
    Dim rowArray As Variant

    Dim ratio1 As Long
    Dim ratio3 As Long
    Dim ratio4 As Long

    Dim i As Long
    Dim cnt As Long
    Dim impose As Boolean

    Dim rng As Variant

    For Each rowArray In moduleMatrix
        Set ratio3Ranges = GetRatio3Ranges(rowArray)

        For Each rng In ratio3Ranges
            ratio3 = rng(1) + 1 - rng(0)
            ratio1 = ratio3 \ 3
            ratio4 = ratio1 * 4
            impose = False

            i = rng(0) - 1

            ' light ratio 1
            cnt = 0
            Do While i >= 0
                If rowArray(i) <= 0 Then
                    cnt = cnt + 1
                    i = i - 1
                Else
                    Exit Do
                End If
            Loop

            If cnt <> ratio1 Then GoTo Continue

            ' dark ratio 1
            cnt = 0
            Do While i >= 0
                If rowArray(i) > 0 Then
                    cnt = cnt + 1
                    i = i - 1
                Else
                    Exit Do
                End If
            Loop

            If cnt <> ratio1 Then GoTo Continue

            ' light ratio 4
            cnt = 0
            Do While i >= 0
                If rowArray(i) <= 0 Then
                    cnt = cnt + 1
                    i = i - 1
                Else
                    Exit Do
                End If
            Loop

            If cnt >= ratio4 Then
                impose = True
            End If

            i = rng(1) + 1

            ' light ratio 1
            cnt = 0
            Do While i <= UBound(rowArray)
                If rowArray(i) <= 0 Then
                    cnt = cnt + 1
                    i = i + 1
                Else
                    Exit Do
                End If
            Loop

            If cnt <> ratio1 Then GoTo Continue

            ' dark ratio 1
            cnt = 0
            Do While i <= UBound(rowArray)
                If rowArray(i) > 0 Then
                    cnt = cnt + 1
                    i = i + 1
                Else
                    Exit Do
                End If
            Loop

            If cnt <> ratio1 Then GoTo Continue

            ' light ratio 4
            cnt = 0
            Do While i <= UBound(rowArray)
                If rowArray(i) <= 0 Then
                    cnt = cnt + 1
                    i = i + 1
                Else
                    Exit Do
                End If
            Loop

            If cnt >= ratio4 Then
                impose = True
            End If

            If impose Then
                penalty = penalty + 40
            End If
Continue:
        Next
    Next

    CalcModuleRatioInRow = penalty

End Function

Private Function GetRatio3Ranges(ByRef arg As Variant) As Collection

    Dim ret As Collection
    Set ret = New Collection

    Dim s As Long
    Dim e As Long

    Dim i As Long

    For i = QuietZone.QUIET_ZONE_WIDTH To UBound(arg) - QuietZone.QUIET_ZONE_WIDTH
        If arg(i) > 0 And arg(i - 1) <= 0 Then
            s = i
        End If

        If arg(i) > 0 And arg(i + 1) <= 0 Then
            e = i

            If (e + 1 - s) Mod 3 = 0 Then
                Call ret.Add(Array(s, e))
            End If
        End If
    Next

    Set GetRatio3Ranges = ret

End Function


'------------------------------------------------------------------------------
' (概要)
'  全体に対する暗モジュールの占める割合について失点を計算します。
'------------------------------------------------------------------------------
Private Function CalcProportionOfDarkModules(ByRef moduleMatrix() As Variant) As Long

    Dim darkCount As Long

    Dim rowArray As Variant
    Dim v As Variant

    For Each rowArray In moduleMatrix
        For Each v In rowArray
            If v > 0 Then
                darkCount = darkCount + 1
            End If
        Next
    Next

    Dim numModules As Double
    numModules = (UBound(moduleMatrix) + 1) ^ 2

    Dim temp As Long
    temp = Int((darkCount / numModules * 100) + 1)
    temp = Abs(temp - 50)
    temp = (temp + 4) \ 5

    CalcProportionOfDarkModules = temp * 10

End Function

