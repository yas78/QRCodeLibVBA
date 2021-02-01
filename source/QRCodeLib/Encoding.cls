VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Encoding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const adTypeBinary As Long = 1
Private Const adTypeText   As Long = 2

Private Const adModeReadWrite As Long = 3

Private m_sr As Object ' Is ADODB.Stream
Private m_charset As String

Public Sub Init(ByVal charsetName As String)
    m_charset = charsetName

    Set m_sr = CreateObject("ADODB.Stream")
    m_sr.Mode = adModeReadWrite
    m_sr.Charset = charsetName
End Sub

Public Property Get Charset() As String
    Charset = m_charset
End Property

Public Function GetBytes(ByVal s As String) As Byte()
    Call m_sr.Open

    m_sr.Position = 0
    m_sr.Type = adTypeText
    Call m_sr.WriteText(s)

    m_sr.Position = 0
    m_sr.Type = adTypeBinary

    Select Case LCase(m_charset)
        Case "utf-16", "unicode"
            m_sr.Position = 2
        Case "utf-8"
            m_sr.Position = 3
        Case Else
            ' NOP
    End Select

    Dim ret() As Byte
    ret = m_sr.Read()
    Call m_sr.Close

    GetBytes = ret
End Function
