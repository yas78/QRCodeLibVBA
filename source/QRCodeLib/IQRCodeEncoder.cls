VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IQRCodeEncoder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ---------------------------------------------------------------------------
' エンコーダーインターフェイス
' ---------------------------------------------------------------------------
Option Explicit

Public Property Get EncodingMode() As EncodingMode: End Property
Public Property Get ModeIndicator() As Long: End Property
Public Property Get CharCount() As Long: End Property
Public Property Get BitCount() As Long: End Property

Public Function Append(ByVal c As String) As Long: End Function
Public Function GetCodewordBitLength(ByVal c As String) As Long: End Function
Public Function GetBytes() As Byte(): End Function
Public Function IsInSubset(ByVal c As String) As Boolean: End Function
Public Function IsInExclusiveSubset(ByVal c As String) As Boolean: End Function
