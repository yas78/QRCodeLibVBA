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
    Dim dr As Direction

    Dim p As Point

    Dim r As Long
    Dim c As Long
    For r = 0 To UBound(img) - 1
        For c = 0 To UBound(img(r)) - 1
            If img(r)(c) = MAX_VALUE Then GoTo Continue
            If Not (img(r)(c) > 0 And img(r)(c + 1) <= 0) Then GoTo Continue

            img(r)(c) = MAX_VALUE
            Set st = New Point
            Call st.Init(c, r)
            Set ptList = New List
            Call ptList.Add(st)

            dr = Direction.Up
            Set p = st.Clone()
            p.Y = p.Y - 1

            Do Until p.Equals(st)
                Select Case dr
                    Case Direction.Up
                        If img(p.Y)(p.X) > 0 Then
                            img(p.Y)(p.X) = MAX_VALUE

                            If img(p.Y)(p.X + 1) <= 0 Then
                                Set p = p.Clone()
                                p.Y = p.Y - 1
                            Else
                                Call ptList.Add(p)
                                dr = Direction.Right
                                Set p = p.Clone()
                                p.X = p.X + 1
                            End If
                        Else
                            Set p = p.Clone()
                            p.Y = p.Y + 1
                            Call ptList.Add(p)
                            dr = Direction.Left
                            Set p = p.Clone()
                            p.X = p.X - 1
                        End If

                    Case Direction.Down
                        If img(p.Y)(p.X) > 0 Then
                            img(p.Y)(p.X) = MAX_VALUE

                            If img(p.Y)(p.X - 1) <= 0 Then
                                Set p = p.Clone()
                                p.Y = p.Y + 1
                            Else
                                Call ptList.Add(p)
                                dr = Direction.Left
                                Set p = p.Clone()
                                p.X = p.X - 1
                            End If
                        Else
                            Set p = p.Clone()
                            p.Y = p.Y - 1
                            Call ptList.Add(p)
                            dr = Direction.Right
                            Set p = p.Clone()
                            p.X = p.X + 1
                        End If

                    Case Direction.Left
                        If img(p.Y)(p.X) > 0 Then
                            img(p.Y)(p.X) = MAX_VALUE

                            If img(p.Y - 1)(p.X) <= 0 Then
                                Set p = p.Clone()
                                p.X = p.X - 1
                            Else
                                Call ptList.Add(p)
                                dr = Direction.Up
                                Set p = p.Clone()
                                p.Y = p.Y - 1
                            End If
                        Else
                            Set p = p.Clone()
                            p.X = p.X + 1
                            Call ptList.Add(p)
                            dr = Direction.Down
                            Set p = p.Clone()
                            p.Y = p.Y + 1
                        End If

                    Case Direction.Right
                        If img(p.Y)(p.X) > 0 Then
                            img(p.Y)(p.X) = MAX_VALUE

                            If img(p.Y + 1)(p.X) <= 0 Then
                                Set p = p.Clone()
                                p.X = p.X + 1
                            Else
                                Call ptList.Add(p)
                                dr = Direction.Down
                                Set p = p.Clone()
                                p.Y = p.Y + 1
                            End If
                        Else
                            Set p = p.Clone()
                            p.X = p.X - 1
                            Call ptList.Add(p)
                            dr = Direction.Up
                            Set p = p.Clone()
                            p.Y = p.Y - 1
                        End If
                Case Else
                    Call Err.Raise(51)
                End Select
            Loop

            Call ptsList.Add(ptList.Items())
Continue:
        Next
    Next

    FindContours = ptsList.Items()
End Function
