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
        Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32" (ByRef lpPictDesc As PICTDESC, ByRef riid As UUID, ByVal fOwn As Boolean, ByRef ppvObj As Any) As Long
    #Else
        Private Declare PtrSafe Function OleLoadPicture Lib "olepro32" (ByVal pStream As IUnknown, ByVal lSize As Long, ByVal fRunmode As Long, ByRef riid As UUID, ByRef ppvObj As Any) As Long
        Private Declare PtrSafe Function OleCreatePictureIndirect Lib "olepro32" (ByRef lpPictDesc As PICTDESC, ByRef riid As UUID, ByVal fOwn As Boolean, ByRef ppvObj As Any) As Long
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
    Private Declare Function OleCreatePictureIndirect Lib "olepro32" (ByRef lpPictDesc As PICTDESC, ByRef riid As UUID, ByVal fOwn As Boolean, ByRef ppvObj As Any) As Long
#End If

Private Const IID_IPictureDisp As String = "{7BF80980-BF32-101A-8BBB-00AA00300CAB}"

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

Private Type PICTDESC
    cbSizeofstruct As Long
    picType        As Long
#If VBA7 Then
    hEmf As LongPtr
#Else
    hEmf As Long
#End If
End Type

Private Const PICTYPE_UNINITIALIZED As Long = (-1)
Private Const PICTYPE_NONE          As Long = 0
Private Const PICTYPE_BITMAP        As Long = 1
Private Const PICTYPE_METAFILE      As Long = 2
Private Const PICTYPE_ICON          As Long = 3
Private Const PICTYPE_ENHMETAFILE   As Long = 4

Public Function ConvertFromDib(ByRef dibData() As Byte) As stdole.IPictureDisp
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
    Call IIDFromString(StrPtr(IID_IPictureDisp), iid)

    Dim ret As stdole.IPictureDisp
    Call OleLoadPicture(stm, sz, WIN32_FALSE, iid, ret)
    Call GlobalFree(hMem)

    Set ConvertFromDib = ret
End Function

#If VBA7 Then
Public Function ConvertFromEmf(ByVal hEmf As LongPtr) As stdole.IPictureDisp
#Else
Public Function ConvertFromEmf(ByVal hEmf As Long) As stdole.IPictureDisp
#End If
    Dim lpPictDesc As PICTDESC

    With lpPictDesc
      .cbSizeofstruct = Len(lpPictDesc)
      .picType = PICTYPE_ENHMETAFILE
      .hEmf = hEmf
    End With

    Dim iid As UUID
    Call IIDFromString(StrPtr(IID_IPictureDisp), iid)

    Dim ret As stdole.IPictureDisp
    Call OleCreatePictureIndirect(lpPictDesc, iid, False, ret)

    Set ConvertFromEmf = ret
End Function
