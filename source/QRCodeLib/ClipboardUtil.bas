Attribute VB_Name = "ClipboardUtil"
Option Private Module
Option Explicit

Private Const GMEM_MOVEABLE As Long = &H2
Private Const GMEM_ZEROINIT As Long = &H40
Private Const GHND          As Long = GMEM_MOVEABLE Or GMEM_ZEROINIT

Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long

Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal pDest As Long, ByVal pSrc As Long, ByVal sz As Long)

Private Declare Function OpenClipboard Lib "user32" (ByVal Hwnd As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long

Const CF_DIB As Long = 8

Public Sub SetDIB(ByRef dibData() As Byte)
    
    Dim sz As Long
    sz = UBound(dibData) - 14 + 1

    Dim hMem As Long
    hMem = GlobalAlloc(GHND, sz)
    If hMem = 0 Then Exit Sub

    Dim lpMem As Long
    lpMem = GlobalLock(hMem)
    If lpMem = 0 Then Exit Sub

    Call MoveMemory(lpMem, VarPtr(dibData(14)), sz)
    Call GlobalUnlock(hMem)

    Call OpenClipboard(0)
    Call EmptyClipboard
    Call SetClipboardData(CF_DIB, lpMem)
    Call CloseClipboard
    
End Sub
