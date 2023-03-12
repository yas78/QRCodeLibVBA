Attribute VB_Name = "Values"
Option Private Module
Option Explicit

Public Const BLANK         As Long = 0
Public Const WORD          As Long = 1
Public Const ALIGNMENT_PTN As Long = 2
Public Const FINDER_PTN    As Long = 3
Public Const FORMAT_INFO   As Long = 4
Public Const SEPARATOR_PTN As Long = 5
Public Const TIMING_PTN    As Long = 6
Public Const VERSION_INFO  As Long = 7

Public Function IsDark(ByVal arg As Long) As Boolean
    IsDark = arg > BLANK
End Function
