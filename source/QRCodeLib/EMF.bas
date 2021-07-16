Attribute VB_Name = "EMF"
Option Private Module
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hDC As LongPtr) As Long
    Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long

    Private Declare PtrSafe Function CreateEnhMetaFile Lib "gdi32" Alias "CreateEnhMetaFileA" (ByVal hDC As LongPtr, ByVal lpFilename As String, ByRef lprc As Any, ByVal lpDesc As String) As LongPtr
    Private Declare PtrSafe Function CloseEnhMetaFile Lib "gdi32" (ByVal hDC As LongPtr) As LongPtr

    Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hDC As LongPtr, ByVal hObject As LongPtr) As LongPtr
    Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long

    Private Declare PtrSafe Function BeginPath Lib "gdi32" (ByVal hDC As LongPtr) As Long
    Private Declare PtrSafe Function EndPath Lib "gdi32" (ByVal hDC As LongPtr) As Long

    Private Declare PtrSafe Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As LongPtr
    Private Declare PtrSafe Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As LongPtr

    Private Declare PtrSafe Function Polygon Lib "gdi32" (ByVal hDC As LongPtr, ByRef lpPoint As POINTAPI, ByVal nCount As Long) As Long

    Private Declare PtrSafe Function SetPolyFillMode Lib "gdi32" (ByVal hDC As LongPtr, ByVal nPolyFillMode As Long) As Long
    Private Declare PtrSafe Function StrokeAndFillPath Lib "gdi32" (ByVal hDC As LongPtr) As Long
#Else
    Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
    Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long

    Private Declare Function CreateEnhMetaFile Lib "gdi32" Alias "CreateEnhMetaFileA" (ByVal hDC As Long, ByVal lpFilename As String, ByRef lprc As RECT, ByVal lpDesc As String) As Long
    Private Declare Function CloseEnhMetaFile Lib "gdi32" (ByVal hDC As Long) As Long

    Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
    Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

    Private Declare Function BeginPath Lib "gdi32" (ByVal hDC As Long) As Long
    Private Declare Function EndPath Lib "gdi32" (ByVal hDC As Long) As Long

    Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
    Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

    Private Declare Function Polygon Lib "gdi32" (ByVal hDC As Long, ByRef lpPoint As POINTAPI, ByVal nCount As Long) As Long

    Private Declare Function SetPolyFillMode Lib "gdi32" (ByVal hDC As Long, ByVal nPolyFillMode As Long) As Long
    Private Declare Function StrokeAndFillPath Lib "gdi32" (ByVal hDC As Long) As Long
#End If

Private Enum PenStyle
    PS_COSMETIC = &H0
    PS_ENDCAP_ROUND = &H0
    PS_JOIN_ROUND = &H0
    PS_SOLID = &H0
    PS_DASH = &H1
    PS_DOT = &H2
    PS_DASHDOT = &H3
    PS_DASHDOTDOT = &H4
    PS_NULL = &H5
    PS_INSIDEFRAME = &H6
    PS_USERSTYLE = &H7
    PS_ALTERNATE = &H8
    PS_ENDCAP_SQUARE = &H100
    PS_ENDCAP_FLAT = &H200
    PS_JOIN_BEVEL = &H1000
    PS_JOIN_MITER = &H2000
    PS_GEOMETRIC = &H10000
End Enum

Private Enum PolygonFillMode
    ALTERNATE = &H1
    WINDING = &H2
End Enum

Private Type RECT
  Left   As Long
  Top    As Long
  Right  As Long
  Bottom As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Const HORZSIZE As Long = 4
Private Const VERTSIZE As Long = 6
Private Const HORZRES  As Long = 8
Private Const VERTRES  As Long = 10

#If VBA7 Then
Public Function GetEMF(ByRef gpPaths() As Variant, _
                       ByVal pictWidth As Long, _
                       ByVal pictHeight As Long, _
                       ByVal foreColorRgb As Long) As LongPtr
#Else
Public Function GetEMF(ByRef gpPaths() As Variant, _
                       ByVal pictWidth As Long, _
                       ByVal pictHeight As Long, _
                       ByVal foreColorRgb As Long) As Long
#End If

#If VBA7 Then
    Dim hScreenDC As LongPtr
#Else
    Dim hScreenDC As Long
#End If

    hScreenDC = GetDC(0)

    Dim mmPerPixelH As Double
    mmPerPixelH = GetDeviceCaps(hScreenDC, HORZSIZE) / GetDeviceCaps(hScreenDC, HORZRES)
    Dim mmPerPixelV As Double
    mmPerPixelV = GetDeviceCaps(hScreenDC, VERTSIZE) / GetDeviceCaps(hScreenDC, VERTRES)

    Call ReleaseDC(0, hScreenDC)

    Dim r As RECT
    With r
        .Left = 0
        .Top = 0
        .Right = mmPerPixelH * pictWidth * 100
        .Bottom = mmPerPixelV * pictHeight * 100
    End With

#If VBA7 Then
    Dim hDC As LongPtr
#Else
    Dim hDC As Long
#End If

    hDC = CreateEnhMetaFile(0, vbNullString, r, vbNullString)

    Call BeginPath(hDC)

    Dim gpPath As Variant
    Dim i As Long
    Dim pArray() As POINTAPI
    For Each gpPath In gpPaths
        ReDim pArray(UBound(gpPath))
        For i = 0 To UBound(gpPath)
            pArray(i).X = gpPath(i).X
            pArray(i).Y = gpPath(i).Y
        Next
        Call Polygon(hDC, pArray(0), UBound(pArray) + 1)
    Next

    Call EndPath(hDC)

#If VBA7 Then
    Dim hBrush    As LongPtr
    Dim hOldBrush As LongPtr
    Dim hPen      As LongPtr
    Dim hOldPen   As LongPtr
#Else
    Dim hBrush    As Long
    Dim hOldBrush As Long
    Dim hPen      As Long
    Dim hOldPen   As Long
#End If

    hBrush = CreateSolidBrush(foreColorRgb)
    hOldBrush = SelectObject(hDC, hBrush)
    hPen = CreatePen(PS_SOLID, 0, foreColorRgb)
    hOldPen = SelectObject(hDC, hPen)

    Call SetPolyFillMode(hDC, PolygonFillMode.ALTERNATE)
    Call StrokeAndFillPath(hDC)

    Call SelectObject(hDC, hOldBrush)
    Call DeleteObject(hBrush)
    Call SelectObject(hDC, hOldPen)
    Call DeleteObject(hPen)

    GetEMF = CloseEnhMetaFile(hDC)
End Function
