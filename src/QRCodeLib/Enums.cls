VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Enums"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum ErrorCorrectionLevel
    L
    M
    Q
    H
End Enum

Public Enum ImageFormat
    fmtBMP = &H10
    fmtPNG = &H20
    fmtSVG = &H30
    fmtEMF = &H40
    fmtTIFF = &H50
    fmtGIF = &H60
    fmtMonochrome = 0
    fmtTrueColor = 1
    fmtBilevel = 2
End Enum

Public Enum OlePicType
    Bitmap = 1
    EnhMetaFile = 4
End Enum

Public Enum BackStyle
    bkTransparent = 0
    bkOpaque = 1
End Enum
