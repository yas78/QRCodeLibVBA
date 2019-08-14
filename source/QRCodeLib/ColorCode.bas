Attribute VB_Name = "ColorCode"
'Option Private Module
Option Explicit


Public Const WHITE As String = "#FFFFFF"
Public Const BLACK As String = "#000000"

Public Function ToRGB(ByVal arg As String) As Long
    If Not Valid(arg) Then Call Err.Raise(5)

    Dim ret As Long
    ret = RGB(CInt("&h" & Mid$(arg, 2, 2)), _
              CInt("&h" & Mid$(arg, 4, 2)), _
              CInt("&h" & Mid$(arg, 6, 2)))

    ToRGB = ret
End Function

Public Function Valid(ByVal arg As String) As Boolean
    Const COLOR_CODE_PTN As String = "[#]" & _
                                     "[0-9A-Fa-f]" & _
                                     "[0-9A-Fa-f]" & _
                                     "[0-9A-Fa-f]" & _
                                     "[0-9A-Fa-f]" & _
                                     "[0-9A-Fa-f]" & _
                                     "[0-9A-Fa-f]"

    Valid = arg Like COLOR_CODE_PTN
End Function
