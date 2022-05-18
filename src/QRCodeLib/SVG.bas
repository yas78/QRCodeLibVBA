Attribute VB_Name = "SVG"
Option Private Module
Option Explicit

Public Function GetSVG(ByRef pts() As Variant, _
                       ByVal pictWidth As Long, _
                       ByVal pictHeight As Long, _
                       ByVal foreColorRgb As Long) As String
    Dim bytes() As Byte
    bytes = BitConverter.GetBytes(foreColorRgb)

    Dim r As String
    Dim g As String
    Dim b As String
    r = Right$("0" & Hex$(bytes(0)), 2)
    g = Right$("0" & Hex$(bytes(1)), 2)
    b = Right$("0" & Hex$(bytes(2)), 2)

    Dim foreRgb As String
    foreRgb = "#" & r & g & b

    Dim buf As New List

    Dim indent As String
    indent = String$(5, " ")

    Dim ptArray As Variant
    Dim i As Long
    For Each ptArray In pts
        Call buf.Add(indent & "M ")

        For i = 0 To UBound(ptArray)
            Call buf.Add(CStr(ptArray(i).X) & "," & CStr(ptArray(i).Y) & " ")
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
