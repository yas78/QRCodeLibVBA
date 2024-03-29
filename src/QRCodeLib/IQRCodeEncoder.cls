VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IQRCodeEncoder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get EncodingMode() As EncodingMode: End Property
Public Property Get ModeIndicator() As Long: End Property
Public Property Get CharCount() As Long: End Property
Public Property Get BitCount() As Long: End Property

Public Sub Init(ByVal charEncoding As Encoding): End Sub
Public Sub Append(ByVal c As String): End Sub
Public Function GetCodewordBitLength(ByVal c As String) As Long: End Function
Public Function GetBytes() As Byte(): End Function
Public Function InSubset(ByVal c As String) As Boolean: End Function
Public Function InExclusiveSubset(ByVal c As String) As Boolean: End Function
