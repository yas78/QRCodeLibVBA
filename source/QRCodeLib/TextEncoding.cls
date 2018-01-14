VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextEncoding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const adTypeBinary      As Long = 1
Private Const adTypeText        As Long = 2

Private Const adModeReadWrite   As Long = 3

Private m_sr        As Object ' Is ADODB.Stream
Private m_charset   As String

Public Property Get Charset() As String

    Charset = m_charset
    
End Property

Public Sub Init(ByVal charsetName As String)
    
    m_charset = charsetName
        
    Set m_sr = CreateObject("ADODB.Stream")
    
    m_sr.Mode = adModeReadWrite
    m_sr.Charset = charsetName
    
End Sub

Public Function GetBytes(ByVal s As String) As Byte()

    Const BOM_LENGTH As Long = 3

    Call m_sr.Open

    m_sr.Position = 0
    m_sr.Type = adTypeText
    Call m_sr.WriteText(s)

    m_sr.Position = 0
    m_sr.Type = adTypeBinary

    Dim buf() As Byte
    buf = m_sr.Read()

    Dim temp() As Byte
    Dim srcIdx  As Long
    Dim destIdx As Long
    
    Select Case LCase(m_charset)
        Case "utf-8"
            ReDim temp(UBound(buf) - BOM_LENGTH)
    
            ' Remove BOM
            For srcIdx = BOM_LENGTH To UBound(buf)
                temp(destIdx) = buf(srcIdx)
                destIdx = destIdx + 1
            Next
            
            GetBytes = temp

        Case Else
            GetBytes = buf
            
    End Select
    
    Call m_sr.Close
    
End Function
