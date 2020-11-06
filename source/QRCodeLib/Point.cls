VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Point"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_x As Long
Private m_y As Long

Public Sub Init(ByVal x As Long, ByVal y As Long)
    m_x = x
    m_y = y
End Sub

Public Property Get x() As Long
    x = m_x
End Property
Public Property Let x(ByVal Value As Long)
    m_x = Value
End Property

Public Property Get y() As Long
    y = m_y
End Property
Public Property Let y(ByVal Value As Long)
    m_y = Value
End Property

Public Function Clone() As Point
    Dim ret As Point
    Set ret = New Point
    Call ret.Init(x, y)
    
    Set Clone = ret
End Function

Public Function Equals(ByVal obj As Point) As Boolean
    Equals = (x = obj.x) And (y = obj.y)
End Function
