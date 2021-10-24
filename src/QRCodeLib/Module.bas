Attribute VB_Name = "Module"
Option Private Module
Option Explicit

Public Function GetNumModulesPerSide(ByVal ver As Long) As Long
    GetNumModulesPerSide = 17 + 4 * ver
End Function
