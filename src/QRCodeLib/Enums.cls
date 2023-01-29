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
    fmtBmp = &H10
    fmtPng = &H20
    fmtSvg = &H30
    fmtEmf = &H40
    fmtTiff = &H50
    fmtGif = &H60
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
