Attribute VB_Name = "Figure"
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

    Dim ptsList As New List
    Dim ptList  As List

    Dim st As Point
    Dim p  As Point
    Dim dr As Direction

    Dim r As Long
    Dim c As Long
    For r = 1 To UBound(img) - 1
        For c = 1 To UBound(img(r)) - 1
            If img(r)(c) = MAX_VALUE Then GoTo Continue
            If img(r)(c) <= 0 Then GoTo Continue
            If img(r)(c - 1) > 0 Then GoTo Continue

            Set st = New Point
            Call st.Init(c, r)
            Set ptList = New List
            Call ptList.Add(st)

            dr = Direction.Down
            Set p = st.Clone()

            Do
                img(p.Y)(p.X) = MAX_VALUE

                Select Case dr
                    Case Direction.Up
                        If img(p.Y - 1)(p.X) <= 0 Then
                            Call ptList.Add(p.Clone())
                            dr = Direction.Left
                            p.X = p.X - 1
                        Else
                            p.Y = p.Y - 1
                            If img(p.Y)(p.X + 1) > 0 Then
                                Call ptList.Add(p.Clone())
                                dr = Direction.Right
                                p.X = p.X + 1
                            End If
                        End If
                    Case Direction.Down
                        If img(p.Y + 1)(p.X) <= 0 Then
                            Call ptList.Add(p.Clone())
                            dr = Direction.Right
                            p.X = p.X + 1
                        Else
                            p.Y = p.Y + 1
                            If img(p.Y)(p.X - 1) > 0 Then
                                Call ptList.Add(p.Clone())
                                dr = Direction.Left
                                p.X = p.X - 1
                            End If
                        End If
                    Case Direction.Left
                        If img(p.Y)(p.X - 1) <= 0 Then
                            Call ptList.Add(p.Clone())
                            dr = Direction.Down
                            p.Y = p.Y + 1
                        Else
                            p.X = p.X - 1
                            If img(p.Y - 1)(p.X) > 0 Then
                                Call ptList.Add(p.Clone())
                                dr = Direction.Up
                                p.Y = p.Y - 1
                            End If
                        End If
                    Case Direction.Right
                        If img(p.Y)(p.X + 1) <= 0 Then
                            Call ptList.Add(p.Clone())
                            dr = Direction.Up
                            p.Y = p.Y - 1
                        Else
                            p.X = p.X + 1
                            If img(p.Y + 1)(p.X) > 0 Then
                                Call ptList.Add(p.Clone())
                                dr = Direction.Down
                                p.Y = p.Y + 1
                            End If
                        End If
                    Case Else
                        Call Err.Raise(51)
                End Select
            Loop Until p.Equals(st)

            Call ptsList.Add(ptList.Items())
Continue:
        Next
    Next

    FindContours = ptsList.Items()
End Function
