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
    penalty = penalty + CalcAdjacentModulesInRowInSameColor(MatrixRotate90(moduleMatrix))

    CalcAdjacentModulesInSameColor = penalty

End Function

'------------------------------------------------------------------------------
' (概要)
'  行の同色隣接モジュールパターンの失点を計算します。
'------------------------------------------------------------------------------
Private Function CalcAdjacentModulesInRowInSameColor(ByRef moduleMatrix() As Variant) As Long

    Dim penalty As Long
    penalty = 0

    Dim r As Long
    Dim c As Long
    Dim cnt As Long
    
    For r = 0 To UBound(moduleMatrix)
        cnt = 1

        For c = 0 To UBound(moduleMatrix(r)) - 1
            If (moduleMatrix(r)(c) > 0) = (moduleMatrix(r)(c + 1) > 0) Then
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
            isSameColor = True
            
            isSameColor = isSameColor And (moduleMatrix(r + 0)(c + 1) > 0 = temp)
            isSameColor = isSameColor And (moduleMatrix(r + 1)(c + 0) > 0 = temp)
            isSameColor = isSameColor And (moduleMatrix(r + 1)(c + 1) > 0 = temp)
    
            If isSameColor Then
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
    penalty = penalty + CalcModuleRatioInRow(MatrixRotate90(moduleMatrixTemp))
    
    CalcModuleRatio = penalty

End Function


'------------------------------------------------------------------------------
' (概要)
'  行の1 : 1 : 3 : 1 : 1 比率のパターンを評価し、失点を返します。
'------------------------------------------------------------------------------
Private Function CalcModuleRatioInRow(ByRef moduleMatrix() As Variant) As Long

    Dim penalty As Long
    penalty = 0

    Dim r As Long
    Dim c As Long
    Dim cols() As Long
    Dim startIndexes  As Collection
    
    Dim i        As Long
    Dim idx      As Long
    Dim modRatio As ModuleRatio
    
    For r = 0 To UBound(moduleMatrix)
        cols = moduleMatrix(r)
        Set startIndexes = New Collection

        Call startIndexes.Add(0)

        For c = 0 To UBound(cols) - 2
            If cols(c) > 0 And cols(c + 1) <= 0 Then
                Call startIndexes.Add(c + 1)
            End If
        Next

        For i = 1 To startIndexes.Count
            idx = startIndexes(i)
            Set modRatio = New ModuleRatio

            Do While idx <= UBound(cols)
                If cols(idx) > 0 Then Exit Do
                modRatio.PreLightRatio4 = modRatio.PreLightRatio4 + 1
                idx = idx + 1
            Loop

            Do While idx <= UBound(cols)
                If cols(idx) <= 0 Then Exit Do
                modRatio.PreDarkRatio1 = modRatio.PreDarkRatio1 + 1
                idx = idx + 1
            Loop

            Do While idx <= UBound(cols)
                If cols(idx) > 0 Then Exit Do
                modRatio.PreLightRatio1 = modRatio.PreLightRatio1 + 1
                idx = idx + 1
            Loop

            Do While idx <= UBound(cols)
                If cols(idx) <= 0 Then Exit Do
                modRatio.CenterDarkRatio3 = modRatio.CenterDarkRatio3 + 1
                idx = idx + 1
            Loop

            Do While idx <= UBound(cols)
                If cols(idx) > 0 Then Exit Do
                modRatio.FolLightRatio1 = modRatio.FolLightRatio1 + 1
                idx = idx + 1
            Loop

            Do While idx <= UBound(cols)
                If cols(idx) <= 0 Then Exit Do
                modRatio.FolDarkRatio1 = modRatio.FolDarkRatio1 + 1
                idx = idx + 1
            Loop

            Do While idx <= UBound(cols)
                If cols(idx) > 0 Then Exit Do
                modRatio.FolLightRatio4 = modRatio.FolLightRatio4 + 1
                idx = idx + 1
            Loop

            If modRatio.PenaltyImposed() Then
                penalty = penalty + 40
            End If

        Next
    Next

    CalcModuleRatioInRow = penalty
    
End Function

'------------------------------------------------------------------------------
' (概要)
'  全体に対する暗モジュールの占める割合について失点を計算します。
'------------------------------------------------------------------------------
Private Function CalcProportionOfDarkModules(ByRef moduleMatrix() As Variant) As Long

    Dim darkCount As Long

    Dim r As Long
    Dim c As Long
    
    For r = 0 To UBound(moduleMatrix)
        For c = 0 To UBound(moduleMatrix(r))
            If moduleMatrix(r)(c) > 0 Then
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

'------------------------------------------------------------------------------
' (概要)
'  左に90度回転した配列を返します。
'------------------------------------------------------------------------------
Private Function MatrixRotate90(ByRef arg() As Variant) As Variant()

    Dim ret() As Variant
    ReDim ret(UBound(arg(0)))

    Dim i As Long
    Dim j As Long
    Dim cols() As Long
    
    For i = 0 To UBound(ret)
        ReDim cols(UBound(arg))
        ret(i) = cols
    Next
    
    Dim k As Long
    k = UBound(ret)
    
    For i = 0 To UBound(ret)
        For j = 0 To UBound(ret(i))
            ret(i)(j) = arg(j)(k - i)
        Next
    Next

    MatrixRotate90 = ret

End Function

