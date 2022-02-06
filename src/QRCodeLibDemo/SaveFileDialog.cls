VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SaveFileDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function GetSaveFileName Lib "comdlg32" Alias "GetSaveFileNameA" (ByRef pOpenfilename As OPENFILENAME) As Long
#Else
    Private Declare Function GetSaveFileName Lib "comdlg32" Alias "GetSaveFileNameA" (byRef pOpenfilename As OPENFILENAME) As Long
#End If

#If VBA7 Then
    Private Type OPENFILENAME
        lStructSize       As Long
        hwndOwner         As LongPtr
        hInstance         As LongPtr
        lpstrFilter       As String
        lpstrCustomFilter As String
        nMaxCustFilter    As Long
        nFilterIndex      As Long
        lpstrFile         As String
        nMaxFile          As Long
        lpstrFileTitle    As String
        nMaxFileTitle     As Long
        lpstrInitialDir   As String
        lpstrTitle        As String
        Flags             As Long
        nFileOffset       As Integer
        nFileExtension    As Integer
        lpstrDefExt       As String
        lCustData         As LongPtr
        lpfnHook          As LongPtr
        lpTemplateName    As String
        pvReserved        As LongPtr
        dwReserved        As Long
        FlagsEx           As Long
    End Type
#Else
    Private Type OPENFILENAME
        lStructSize       As Long
        hwndOwner         As Long
        hInstance         As Long
        lpstrFilter       As String
        lpstrCustomFilter As String
        nMaxCustFilter    As Long
        nFilterIndex      As Long
        lpstrFile         As String
        nMaxFile          As Long
        lpstrFileTitle    As String
        nMaxFileTitle     As Long
        lpstrInitialDir   As String
        lpstrTitle        As String
        Flags             As Long
        nFileOffset       As Integer
        nFileExtension    As Integer
        lpstrDefExt       As String
        lCustData         As Long
        lpfnHook          As Long
        lpTemplateName    As String
        pvReserved        As Long
        dwReserved        As Long
        FlagsEx           As Long
  End Type
#End If

Private Const OFN_OVERWRITEPROMPT As Long = &H2
Private Const OFN_PATHMUSTEXIST   As Long = &H800
Private Const OFN_EXPLORER        As Long = &H80000

Private Const BF_SIZE As Long = 256

Private m_lpstrFilter     As String
Private m_nFilterIndex    As Long
Private m_lpstrFile       As String

Private m_lpstrInitialDir As String
Private m_lpstrTitle      As String

Public Property Get Filter() As String
    Filter = m_lpstrFilter
End Property
Public Property Let Filter(ByVal Value As String)
    m_lpstrFilter = Value
End Property

Public Property Get FilterIndex() As Long
    FilterIndex = m_nFilterIndex
End Property
Public Property Let FilterIndex(ByVal Value As Long)
    m_nFilterIndex = Value
End Property

Public Property Get FileName() As String
    FileName = m_lpstrFile
End Property

Public Property Get InitialDirectory() As String
    InitialDirectory = m_lpstrInitialDir
End Property
Public Property Let InitialDirectory(ByVal Value As String)
    m_lpstrInitialDir = Value
End Property

Public Property Get Title() As String
    Title = m_lpstrTitle
End Property
Public Property Let Title(ByVal Value As String)
    m_lpstrTitle = Value
End Property

#If VBA7 Then
Public Function ShowDialog(ByVal hwndOwner As LongPtr) As Boolean
#Else
Public Function ShowDialog(ByVal hwndOwner As Long) As Boolean
#End If
    m_lpstrFile = ""

    Dim ofna As OPENFILENAME
    With ofna
        .lStructSize = LenB(ofna)
        .hInstance = 0
        .hwndOwner = hwndOwner

        If Len(m_lpstrFilter) = 0 Then
            .lpstrFilter = "All files(*.*)" & vbNullChar & "*.*" & vbNullChar & vbNullChar
        Else
            .lpstrFilter = Replace(m_lpstrFilter, ",", vbNullChar) & vbNullChar & vbNullChar
        End If

        .nFilterIndex = m_nFilterIndex
        .lpstrFile = String$(BF_SIZE, vbNullChar)
        .nMaxFile = LenB(.lpstrFile)
        .nMaxFileTitle = 0
        .lpstrInitialDir = m_lpstrInitialDir
        .lpstrTitle = m_lpstrTitle
        .Flags = OFN_EXPLORER Or OFN_PATHMUSTEXIST Or OFN_OVERWRITEPROMPT
    End With

    Dim ret As Boolean
    ret = CBool(GetSaveFileName(ofna))

    If Not ret Then
        ShowDialog = False
        Exit Function
    End If

    With ofna
        m_nFilterIndex = .nFilterIndex
        m_lpstrFile = Left$(.lpstrFile, InStr(.lpstrFile, vbNullChar) - 1)
    End With

    Dim ext As String
    ext = GetExtension(m_nFilterIndex)

    If LCase$(m_lpstrFile) Like ("*." & ext) = False Then
        m_lpstrFile = m_lpstrFile & "." & ext
    End If

    ShowDialog = ret
End Function

Private Function GetExtension(ByVal filterIdx As Long) As String
    If Len(m_lpstrFilter) = 0 Then
        GetExtension = ""
        Exit Function
    End If

    Dim fileFilters As Variant
    fileFilters = Split(m_lpstrFilter, ",")

    Dim ext As String
    ext = fileFilters(filterIdx * 2 - 1)

    Dim token As String
    token = LCase$(Split(ext, ";")(0))

    Dim ret As String
    ret = Mid$(token, InStr(1, token, ".") + 1)
    GetExtension = ret
End Function
