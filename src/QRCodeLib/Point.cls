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

Public Sub Init(ByVal posX As Long, ByVal posY As Long)
    m_x = posX
    m_y = posY
End Sub

Public Property Get X() As Long
    X = m_x
End Property
Public Property Let X(ByVal Value As Long)
    m_x = Value
End Property

Public Property Get Y() As Long
    Y = m_y
End Property
Public Property Let Y(ByVal Value As Long)
    m_y = Value
End Property

Public Function Clone() As Point
    Dim ret As New Point
    Call ret.Init(m_x, m_y)
    Set Clone = ret
End Function

Public Function Equals(ByVal obj As Point) As Boolean
    Equals = (m_x = obj.X) And (m_y = obj.Y)
End Function
