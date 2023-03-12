VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Codeword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_totalNumbers() As Long

Private Sub Class_Initialize()
    m_totalNumbers = LongArray( _
        -1, _
        26, 44, 70, 100, 134, 172, 196, 242, 292, 346, _
        404, 466, 532, 581, 655, 733, 815, 901, 991, 1085, _
        1156, 1258, 1364, 1474, 1588, 1706, 1828, 1921, 2051, 2185, _
        2323, 2465, 2611, 2761, 2876, 3034, 3196, 3362, 3532, 3706 _
    )
End Sub

Public Function GetTotalNumber(ByVal ver As Long) As Long
    GetTotalNumber = m_totalNumbers(ver)
End Function
