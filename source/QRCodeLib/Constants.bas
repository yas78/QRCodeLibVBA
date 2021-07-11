Attribute VB_Name = "Constants"
Option Explicit

Public Const MIN_VERSION As Long = 1
Public Const MAX_VERSION As Long = 40

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
    fmtMonochrome = 0
    fmtTrueColor = 1
End Enum

Public Enum OlePicType
    Bitmap = 1
    EnhMetaFile = 4
End Enum

Public Enum BackStyle
    bkTransparent = 0
    bkOpaque = 1
End Enum
