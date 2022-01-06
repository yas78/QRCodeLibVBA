Attribute VB_Name = "ClipboardUtil"
Option Private Module
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal pDest As LongPtr, ByVal pSrc As LongPtr, ByVal sz As Long)
    Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As LongPtr
    Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
    Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hWnd As LongPtr) As Long
    Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
#Else
    Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal pDest As Long, ByVal pSrc As Long, ByVal sz As Long)
    Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
    Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function CloseClipboard Lib "user32" () As Long
    Private Declare Function EmptyClipboard Lib "user32" () As Long
    Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
#End If

Private Const GMEM_MOVEABLE As Long = &H2
Private Const GMEM_ZEROINIT As Long = &H40
Private Const GHND          As Long = GMEM_MOVEABLE Or GMEM_ZEROINIT

Private Const CF_DIB         As Long = 8
Private Const CF_ENHMETAFILE As Long = 14

Public Sub SetDIB(ByRef dibData() As Byte)
    Dim sz As Long
    sz = UBound(dibData) - 14 + 1

#If VBA7 Then
    Dim hMem As LongPtr
#Else
    Dim hMem As Long
#End If

    hMem = GlobalAlloc(GHND, sz)
    If hMem = 0 Then Exit Sub

#If VBA7 Then
    Dim lpMem As LongPtr
#Else
    Dim lpMem As Long
#End If

    lpMem = GlobalLock(hMem)
    If lpMem = 0 Then Exit Sub

    Call MoveMemory(lpMem, VarPtr(dibData(14)), sz)
    Call GlobalUnlock(hMem)

    Call OpenClipboard(0)
    Call EmptyClipboard
    Call SetClipboardData(CF_DIB, lpMem)
    Call CloseClipboard
End Sub

#If VBA7 Then
Public Sub SetEMF(ByVal hEmf As LongPtr)
#Else
Public Sub SetEMF(ByVal hEmf As Long)
#End If
    Call OpenClipboard(0)
    Call EmptyClipboard
    Call SetClipboardData(CF_ENHMETAFILE, hEmf)
    Call CloseClipboard
End Sub
