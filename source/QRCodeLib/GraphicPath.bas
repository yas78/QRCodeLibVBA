Attribute VB_Name = "GraphicPath"
Option Private Module
Option Explicit

Private Enum Direction
    Up
    Down
    Left
    Right
End Enum

Public Function FindContours(ByRef img() As Variant) As Variant()
    Const MAX_VALUE As Long = &H7FFFFFFF

    Dim gpPaths As New List
    Dim gpPath  As List

    Dim st As Point
    Dim dr As Direction

    Dim x As Long
    Dim y As Long
    Dim p As Point

    For y = 0 To UBound(img) - 1
        For x = 0 To UBound(img(y)) - 1
            If img(y)(x) = MAX_VALUE Then GoTo Continue
            If Not (img(y)(x) > 0 And img(y)(x + 1) <= 0) Then GoTo Continue

            img(y)(x) = MAX_VALUE
            Set st = New Point
            Call st.Init(x, y)
            Set gpPath = New List
            Call gpPath.Add(st)

            dr = Direction.Up
            Set p = st.Clone()
            p.y = p.y - 1

            Do
                Select Case dr
                    Case Direction.Up
                        If img(p.y)(p.x) > 0 Then
                            img(p.y)(p.x) = MAX_VALUE

                            If img(p.y)(p.x + 1) <= 0 Then
                                Set p = p.Clone()
                                p.y = p.y - 1
                            Else
                                Call gpPath.Add(p)
                                dr = Direction.Right
                                Set p = p.Clone()
                                p.x = p.x + 1
                            End If
                        Else
                            Set p = p.Clone()
                            p.y = p.y + 1
                            Call gpPath.Add(p)
                            dr = Direction.Left
                            Set p = p.Clone()
                            p.x = p.x - 1
                        End If

                    Case Direction.Down
                        If img(p.y)(p.x) > 0 Then
                            img(p.y)(p.x) = MAX_VALUE

                            If img(p.y)(p.x - 1) <= 0 Then
                                Set p = p.Clone()
                                p.y = p.y + 1
                            Else
                                Call gpPath.Add(p)
                                dr = Direction.Left
                                Set p = p.Clone()
                                p.x = p.x - 1
                            End If
                        Else
                            Set p = p.Clone()
                            p.y = p.y - 1
                            Call gpPath.Add(p)
                            dr = Direction.Right
                            Set p = p.Clone()
                            p.x = p.x + 1
                        End If

                    Case Direction.Left
                        If img(p.y)(p.x) > 0 Then
                            img(p.y)(p.x) = MAX_VALUE

                            If img(p.y - 1)(p.x) <= 0 Then
                                Set p = p.Clone()
                                p.x = p.x - 1
                            Else
                                Call gpPath.Add(p)
                                dr = Direction.Up
                                Set p = p.Clone()
                                p.y = p.y - 1
                            End If
                        Else
                            Set p = p.Clone()
                            p.x = p.x + 1
                            Call gpPath.Add(p)
                            dr = Direction.Down
                            Set p = p.Clone()
                            p.y = p.y + 1
                        End If

                    Case Direction.Right
                        If img(p.y)(p.x) > 0 Then
                            img(p.y)(p.x) = MAX_VALUE

                            If img(p.y + 1)(p.x) <= 0 Then
                                Set p = p.Clone()
                                p.x = p.x + 1
                            Else
                                Call gpPath.Add(p)
                                dr = Direction.Down
                                Set p = p.Clone()
                                p.y = p.y + 1
                            End If
                        Else
                            Set p = p.Clone()
                            p.x = p.x - 1
                            Call gpPath.Add(p)
                            dr = Direction.Up
                            Set p = p.Clone()
                            p.y = p.y - 1
                        End If
                Case Else
                    Call Err.Raise(51)
                End Select
            Loop While Not p.Equals(st)

            Call gpPaths.Add(gpPath.Items())
Continue:
        Next
    Next

    FindContours = gpPaths.Items()
End Function


