Attribute VB_Name = "SVG"
Option Private Module
Option Explicit

Public Function GetSVG(ByRef gpPaths() As Variant, _
                       ByVal pictWidth As Long, _
                       ByVal pictHeight As Long, _
                       ByVal foreRgb As String) As String
    Dim buf As New List

    Dim indent As String
    indent = String(5, " ")

    Dim gpPath As Variant
    Dim i As Long
    For Each gpPath In gpPaths
        Call buf.Add(indent & "M ")

        For i = 0 To UBound(gpPath)
            Call buf.Add(CStr(gpPath(i).X) & "," & CStr(gpPath(i).Y) & " ")
        Next
        Call buf.Add("Z" & vbNewLine)
    Next

    Dim data As String
    data = Trim$(Join(buf.Items(), ""))
    data = Left$(data, Len(data) - Len(vbNewLine))

    Dim ret As String
    ret = "<svg version=""1.1"" xmlns=""http://www.w3.org/2000/svg"" xmlns:xlink=""http://www.w3.org/1999/xlink""" & vbNewLine & _
          "  width=""" & CStr(pictWidth) & "px"" height=""" & CStr(pictHeight) & "px"" viewBox=""0 0 " & CStr(pictWidth) & " " & CStr(pictHeight) & """>" & vbNewLine & _
          "<path fill=""" & foreRgb & """ stroke=""" & foreRgb & """ stroke-width=""1""" & vbNewLine & _
          "  d=""" & data & """ />" & vbNewLine & _
          "</svg>"

    GetSVG = ret
End Function
