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

Private Const CHARSET_SHIFT_JIS As String = "shift_jis"
Private Const CHARSET_UTF8      As String = "utf-8"

Private Const UTF8_BOM_LENGTH As Long = 3

Private m_sr        As Object ' Is ADODB.Stream
Private m_charset   As String

Public Property Get Charset() As String

    Charset = m_charset
    
End Property

Public Sub Init(ByVal charsetName As String)
    
    If LCase(charsetName) <> CHARSET_SHIFT_JIS And _
       LCase(charsetName) <> CHARSET_UTF8 Then
        Err.Raise 5
    End If
    
    m_charset = LCase(charsetName)
        
    Set m_sr = CreateObject("ADODB.Stream")
    
    m_sr.Mode = adModeReadWrite
    m_sr.Charset = charsetName
    
End Sub

Public Function GetBytes(ByVal s As String) As Byte()

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
    
    Select Case m_charset
        Case CHARSET_UTF8
            ReDim temp(UBound(buf) - UTF8_BOM_LENGTH)
    
            ' Remove BOM
            For srcIdx = UTF8_BOM_LENGTH To UBound(buf)
                temp(destIdx) = buf(srcIdx)
                destIdx = destIdx + 1
            Next
            
            GetBytes = temp

        Case CHARSET_SHIFT_JIS
            GetBytes = buf
            
        Case Else
            Err.Raise 51
            
    End Select
    
    Call m_sr.Close
    
End Function
