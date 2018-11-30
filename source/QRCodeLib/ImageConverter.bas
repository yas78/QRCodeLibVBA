Attribute VB_Name = "ImageConverter"
Option Private Module
Option Explicit

Private Const IID_IPicture As String = "{7BF80980-BF32-101A-8BBB-00AA00300CAB}"

Private Const S_OK As Long = &H0

Private Const GMEM_MOVEABLE As Long = &H2
Private Const GMEM_ZEROINIT As Long = &H40
Private Const GHND          As Long = GMEM_MOVEABLE Or GMEM_ZEROINIT

Private Declare Function IIDFromString Lib "ole32" ( _
    ByVal lpsz As Long, ByRef lpiid As UUID) As Long

Private Declare Function CreateStreamOnHGlobal Lib "ole32" ( _
    ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, _
    ByRef ppstm As Any) As Long
Private Declare Function OleLoadPicture Lib "olepro32" ( _
    ByVal pStream As Long, ByVal lSize As Long, ByVal fRunmode As Long, _
    ByRef riid As UUID, ByRef ppvObj As Any) As Long

Private Declare Function GlobalAlloc Lib "kernel32" ( _
    ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" ( _
    ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" ( _
    ByVal hMem As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" ( _
    ByVal hMem As Long) As Long

Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByVal pDest As Long, ByVal pSrc As Long, ByVal sz As Long)

Private Type UUID
    Data1    As Long
    Data2    As Integer
    Data3    As Integer
    Data4(7) As Byte
End Type

Public Function ConvertFrom(ByRef dibData() As Byte) As IPicture

    Dim sz As Long
    sz = UBound(dibData) + 1

    Dim hMem As Long
    hMem = GlobalAlloc(GMEM_MOVEABLE, sz)
    If hMem = 0 Then Exit Function

    Dim lpMem As Long
    lpMem = GlobalLock(hMem)
    If lpMem = 0 Then Exit Function
    
    Call MoveMemory(lpMem, VarPtr(dibData(0)), sz)
    Call GlobalUnlock(hMem)
    
    Dim stm  As IUnknown
    Dim lRet As Long
    lRet = CreateStreamOnHGlobal(hMem, 1, stm)
    If lRet <> S_OK Then Exit Function

    Dim iid As UUID
    Call IIDFromString(StrPtr(IID_IPicture), iid)
    
    Dim ret As IPicture
    Call OleLoadPicture(ObjPtr(stm), sz, 0, iid, ret)
    Call GlobalFree(hMem)
    
    Set ConvertFrom = ret

End Function
