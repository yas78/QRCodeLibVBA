Attribute VB_Name = "ImageConverter"
Option Private Module
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal pDest As LongPtr, ByVal pSrc As LongPtr, ByVal sz As Long)
    Private Declare PtrSafe Function IIDFromString Lib "ole32" (ByVal lpsz As LongPtr, ByRef lpiid As UUID) As Long
    Private Declare PtrSafe Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As LongPtr, ByVal fDeleteOnRelease As Long, ByRef ppstm As Any) As Long
    Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As LongPtr
    Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
    Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr

    #If Win64 Then
        Private Declare PtrSafe Function OleLoadPicture Lib "oleaut32" (ByVal pStream As IUnknown, ByVal lSize As Long, ByVal fRunmode As Long, ByRef riid As UUID, ByRef ppvObj As Any) As Long
    #Else
        Private Declare PtrSafe Function OleLoadPicture Lib "olepro32" (ByVal pStream As IUnknown, ByVal lSize As Long, ByVal fRunmode As Long, ByRef riid As UUID, ByRef ppvObj As Any) As Long
    #End If
#Else
    Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal pDest As Long, ByVal pSrc As Long, ByVal sz As Long)
    Private Declare Function IIDFromString Lib "ole32" (ByVal lpsz As Long, ByRef lpiid As UUID) As Long
    Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ByRef ppstm As Any) As Long
    Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
    Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function OleLoadPicture Lib "olepro32" (ByVal pStream As IUnknown, ByVal lSize As Long, ByVal fRunmode As Long, ByRef riid As UUID, ByRef ppvObj As Any) As Long
#End If

Private Const IID_IPicture As String = "{7BF80980-BF32-101A-8BBB-00AA00300CAB}"

Private Const S_OK As Long = 0

Private Const GMEM_MOVEABLE As Long = &H2
Private Const GMEM_ZEROINIT As Long = &H40
Private Const GHND          As Long = GMEM_MOVEABLE Or GMEM_ZEROINIT

Private Const WIN32_TRUE  As Long = 1
Private Const WIN32_FALSE As Long = 0

Private Type UUID
    Data1    As Long
    Data2    As Integer
    Data3    As Integer
    Data4(7) As Byte
End Type

Public Function ConvertFrom(ByRef dibData() As Byte) As stdole.IPicture

    Dim sz As Long
    sz = UBound(dibData) + 1

#If VBA7 Then
    Dim hMem As LongPtr
#Else
    Dim hMem As Long
#End If

    hMem = GlobalAlloc(GMEM_MOVEABLE, sz)
    If hMem = 0 Then Exit Function

#If VBA7 Then
    Dim lpMem As LongPtr
#Else
    Dim lpMem As Long
#End If

    lpMem = GlobalLock(hMem)
    If lpMem = 0 Then Exit Function

    Call MoveMemory(lpMem, VarPtr(dibData(0)), sz)
    Call GlobalUnlock(hMem)

    Dim stm  As IUnknown
    Dim lRet As Long
    lRet = CreateStreamOnHGlobal(hMem, WIN32_TRUE, stm)
    If lRet <> S_OK Then Exit Function

    Dim iid As UUID
    Call IIDFromString(StrPtr(IID_IPicture), iid)

    Dim ret As stdole.IPicture
    Call OleLoadPicture(stm, sz, WIN32_FALSE, iid, ret)
    Call GlobalFree(hMem)

    Set ConvertFrom = ret

End Function
