VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_buf As Object

Private Sub Class_Initialize()
    Set m_buf = CreateObject("Scripting.Dictionary")
End Sub

Public Sub Add(ByVal arg As Variant)
    Call m_buf.Add(m_buf.Count, arg)
End Sub

Public Function Items() As Variant()
    Items = m_buf.Items()
End Function
