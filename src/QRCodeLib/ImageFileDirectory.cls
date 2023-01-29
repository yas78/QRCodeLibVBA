VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ImageFileDirectory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_entries As Collection
Private m_nextIFDOffset As Long

Private Sub Class_Initialize()
    Set m_entries = New Collection
    m_nextIFDOffset = 0
End Sub

Public Property Get NextIFDOffset() As Long
    NextIFDOffset = m_nextIFDOffset
End Property
Public Property Let NextIFDOffset(ByVal Value As Long)
    m_nextIFDOffset = Value
End Property

Public Property Get Length() As Long
    Length = 2 + 4 + (12 * m_entries.Count)
End Property

Public Function GetEntries() As IFDEntry()
    Dim ret() As IFDEntry
    ReDim ret(m_entries.Count - 1)

    Dim e As IFDEntry
    Dim idx As Long
    idx = 0
    For Each e In m_entries
        Set ret(idx) = e
        idx = idx + 1
    Next

    GetEntries = ret
End Function

Public Function GetBytes() As Byte()
    Dim buf As New BinaryWriter
    Dim i As Long
    Dim bytes() As Byte

    Call buf.WriteAny(CInt(m_entries.Count))

    Dim e As IFDEntry
    For Each e In m_entries
        Call buf.WriteAny(e.GetBytes())
    Next

    Call buf.WriteAny(m_nextIFDOffset)

    Call buf.Flush(GetBytes)
End Function

Public Function Add(ByVal tag As TagID, ByVal v As Long) As IFDEntry
    Dim ret As New IFDEntry
    Call m_entries.Add(ret)

    ret.TagID = tag
    ret.Value = v
    ret.Count = 1

    Select Case tag
        Case TagID.ImageWidth
            ret.FieldType = TiffFieldType.Long
        Case TagID.ImageLength
            ret.FieldType = TiffFieldType.Long
        Case TagID.BitsPerSample
            ret.FieldType = TiffFieldType.Short
        Case TagID.Compression
            ret.FieldType = TiffFieldType.Short
        Case TagID.PhotometricInterpretation
            ret.FieldType = TiffFieldType.Short
        Case TagID.StripOffsets
            ret.FieldType = TiffFieldType.Long
        Case TagID.SamplesPerPixel
            ret.FieldType = TiffFieldType.Short
        Case TagID.RowsPerStrip
            ret.FieldType = TiffFieldType.Long
        Case TagID.StripByteCounts
            ret.FieldType = TiffFieldType.Long
        Case TagID.XResolution
            ret.FieldType = TiffFieldType.Rational
        Case TagID.YResolution
            ret.FieldType = TiffFieldType.Rational
        Case TagID.ColorMap
            ret.FieldType = TiffFieldType.Short
        Case Else
            Call Err.Raise(5)
    End Select

    Set Add = ret
End Function
