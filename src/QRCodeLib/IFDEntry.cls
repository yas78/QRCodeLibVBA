VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IFDEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_tagID     As Integer
Private m_fieldType As TiffFieldType
Private m_count     As Long
Private m_value     As Long

Public Property Get TagID() As Integer
    TagID = m_tagID
End Property
Public Property Let TagID(ByVal Value As Integer)
    m_tagID = Value
End Property

Public Property Get FieldType() As TiffFieldType
    FieldType = m_fieldType
End Property
Public Property Let FieldType(ByVal Value As TiffFieldType)
    m_fieldType = Value
End Property

Public Property Get Count() As Long
    Count = m_count
End Property
Public Property Let Count(ByVal Value As Long)
    m_count = Value
End Property

Public Property Get Value() As Long
    Value = m_value
End Property
Public Property Let Value(ByVal Value As Long)
    m_value = Value
End Property

Public Function GetDataSize() As Long
    Dim ret As Long

    Select Case m_fieldType
        Case TiffFieldType.[Byte], TiffFieldType.Ascii
            ret = 1 * m_count
        Case TiffFieldType.Short
            ret = 2 * m_count
        Case TiffFieldType.[Long]
            ret = 4 * m_count
        Case TiffFieldType.Rational
            ret = 8 * m_count
        Case Else
            Call Err.Raise(51)
    End Select

    GetDataSize = ret
End Function

Public Function GetBytes() As Byte()
    Dim ret(11) As Byte

    Dim bytes() As Byte
    Dim idx As Long
    idx = 0

    bytes = BitConverter.GetBytes(m_tagID)
    idx = ArrayUtil.CopyAll(ret, idx, bytes)
    bytes = BitConverter.GetBytes(m_fieldType)
    idx = ArrayUtil.Copy(ret, idx, bytes, 0, 2)
    bytes = BitConverter.GetBytes(m_count)
    idx = ArrayUtil.CopyAll(ret, idx, bytes)
    bytes = BitConverter.GetBytes(m_value)
    idx = ArrayUtil.CopyAll(ret, idx, bytes)

    GetBytes = ret
End Function
