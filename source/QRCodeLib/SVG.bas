Attribute VB_Name = "SVG"
Option Private Module
Option Explicit

Public Function BuildSVG(ByRef gpPaths() As Variant, _
                         ByVal pictWidth As Long, _
                         ByVal pictHeight As Long, _
                         ByVal foreRgb As String) As String
    Dim buf As New List
    
    Dim indent As String
    indent = String(11, " ")
    
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
    ret = _
        "<?xml version='1.0' encoding='UTF-8' standalone='no'?>" & vbNewLine & _
        "<!DOCTYPE svg PUBLIC '-//W3C//DTD SVG 20010904//EN'" & vbNewLine & _
        "    'http://www.w3.org/TR/2001/REC-SVG-20010904/DTD/svg10.dtd'>" & vbNewLine & _
        "<svg xmlns='http://www.w3.org/2000/svg'" & vbNewLine & _
        "    width='" & CStr(pictWidth) & "' height='" & CStr(pictHeight) & "' viewBox='0 0 " & CStr(pictWidth) & " " & CStr(pictHeight) & "'>" & vbNewLine & _
        "    <path fill='" & foreRgb & "' stroke='" & foreRgb & "' stroke-width='1'" & vbNewLine & _
        "        d='" & data & "'" & vbNewLine & _
        "    />" & vbNewLine & _
        "</svg>"

    BuildSVG = ret
End Function
