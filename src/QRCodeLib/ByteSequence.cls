VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ByteSequence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const adTypeBinary As Long = 1
Private Const adModeReadWrite As Long = 3

Private m_buf As Object

Private Sub Class_Initialize()
    Call Init
End Sub

Public Sub Init()
    Set m_buf = CreateObject("ADODB.Stream")
    m_buf.Mode = adModeReadWrite
    m_buf.Type = adTypeBinary
    Call m_buf.Open
End Sub

Public Sub Append(ByRef arg As Variant, Optional ByVal reverse As Boolean = False)
    Dim t As VbVarType
    t = VarType(arg)

    Dim i As Long
    Dim j As Long
    Dim bytes() As Byte

    If (t And VbVarType.vbArray) = VbVarType.vbArray Then
        Select Case t - VbVarType.vbArray
            Case VbVarType.vbByte
                Call m_buf.Write(arg)
            Case VbVarType.vbInteger, VbVarType.vbLong
                For i = 0 To UBound(arg)
                    bytes = BitConverter.GetBytes(arg(i), reverse)
                    Call m_buf.Write(bytes)
                Next
#If Win64 Then
            Case VbVarType.vbLongLong
                For i = 0 To UBound(arg)
                    bytes = BitConverter.GetBytes(arg(i), reverse)
                    Call m_buf.Write(bytes)
                Next
#End If
            Case Else
                Call Err.Raise(5)
        End Select
    Else
        Select Case VarType(arg)
            Case VbVarType.vbByte, VbVarType.vbInteger, VbVarType.vbLong
                bytes = BitConverter.GetBytes(arg, reverse)
                Call m_buf.Write(bytes)
#If Win64 Then
            Case VbVarType.vbLongLong
                bytes = BitConverter.GetBytes(arg, reverse)
                Call m_buf.Write(bytes)
#End If
            Case Else
                Call Err.Raise(5)
        End Select
    End If
End Sub

Public Function Flush() As Byte()
    m_buf.Position = 0
    Flush = m_buf.Read()
    Call m_buf.Close
    Call Init
End Function
