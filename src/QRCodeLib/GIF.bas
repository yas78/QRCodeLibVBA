Attribute VB_Name = "GIF"
Option Private Module
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal pDest As LongPtr, ByVal pSrc As LongPtr, ByVal sz As Long)
#Else
    Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal pDest As Long, ByVal pSrc As Long, ByVal sz As Long)
#End If

Private Type GifHeader
    ExtensionIntroducer() As Byte
    Version()             As Byte
End Type

Private Type LogicalScreenDescriptor
    LogicalScreenWidth   As Integer
    LogicalScreenHeight  As Integer
    PackedFields         As Byte
    BackgroundColorIndex As Byte
    PixelAspectRatio     As Byte
End Type

Private Type GraphicControlExtension
    ExtensionIntroducer   As Byte
    GraphicControlLabel   As Byte
    BlockSize             As Byte
    PackedFields          As Byte
    DelayTime             As Integer
    TransparentColorIndex As Byte
    BlockTerminator       As Byte
End Type

Private Type ImageDescriptor
    ImageSeparator     As Byte
    ImageLeftPosition  As Integer
    ImageTopPosition   As Integer
    ImageWidth         As Integer
    ImageHeight        As Integer
    PackedFields       As Byte
    LZWMinimumCodeSize As Byte
End Type

Private Type ImageBlock
    Size        As Byte
    BlockData() As Byte
End Type

Private m_dict      As Object
Private m_buf       As List
Private m_clearCode As String
Private m_endCode   As String
Private m_bitsLen   As Long

Public Function GetGIF(ByRef data() As Byte, _
                       ByVal pictWidth As Long, _
                       ByVal pictHeight As Long, _
                       ByRef palette() As Long, _
                       ByVal bkTransparent As Boolean, _
                       ByVal transparentIndex As Long) As Byte()
    Dim bpp As Long
    bpp = 0

    Dim i As Long
    For i = 1 To 8
        If (UBound(palette) + 1) = 2 ^ i Then
            bpp = i
            Exit For
        End If
    Next

    If bpp = 0 Then Call Err.Raise(5)

    Dim hdr As GifHeader
    Call MakeGifHeader(hdr)

    Dim lsDesc As LogicalScreenDescriptor
    Call MakeLogicalScreenDescriptor(pictWidth, pictHeight, bpp, True, lsDesc)

    Dim gcTable() As Byte
    Call MakeGlobalColorTable(palette, bpp, gcTable)

    Dim gcExt As GraphicControlExtension
    If bkTransparent Then
        Call MakeGraphicControlExtension(bkTransparent, transparentIndex, gcExt)
    End If

    Dim imgDesc As ImageDescriptor
    Call MakeImageDescriptor(pictWidth, pictHeight, bpp, imgDesc)

    Dim imgBlocks() As ImageBlock
    Call MakeImageBlocks(data, bpp, imgBlocks)

    Dim trailer As Byte
    trailer = &H3B

    Dim ret() As Byte
    Call ToBytes(hdr, lsDesc, gcTable, gcExt, imgDesc, imgBlocks, trailer, ret)

    GetGIF = ret
End Function

Private Sub MakeGifHeader(ByRef hdr As GifHeader)
    With hdr
        .ExtensionIntroducer = StrConv("GIF", vbFromUnicode)
        .Version = StrConv("89a", vbFromUnicode)
    End With
End Sub

Private Sub MakeLogicalScreenDescriptor( _
    ByVal pictWidth As Long, _
    ByVal pictHeight As Long, _
    ByVal bpp As Long, _
    ByVal hasGlobalColorTable As Boolean, _
    ByRef lsDesc As LogicalScreenDescriptor)

    Dim globalColorTableFlag As Long
    globalColorTableFlag = IIf(hasGlobalColorTable, 1 * 2 ^ 7, 0)

    Dim colorResolution As Long
    colorResolution = (bpp - 1) * 2 ^ 4

    Dim sortFlag As Long
    sortFlag = 0 * 2 ^ 3

    Dim sizeOfGlobalColorTable As Long
    sizeOfGlobalColorTable = bpp - 1

    With lsDesc
        .LogicalScreenWidth = pictWidth
        .LogicalScreenHeight = pictHeight
        .PackedFields = globalColorTableFlag Or _
                        colorResolution Or _
                        sortFlag Or _
                        sizeOfGlobalColorTable
        .BackgroundColorIndex = 0
        .PixelAspectRatio = 0
    End With
End Sub

Private Sub MakeGlobalColorTable(ByRef palette() As Long, ByVal bpp As Long, gcTable() As Byte)
    ReDim gcTable(2 ^ bpp * 3 - 1)

    Dim idx As Long
    idx = 0

    Dim v As Variant
    For Each v In palette
        gcTable(idx + 0) = CByte(v And &HFF&)
        gcTable(idx + 1) = CByte((v And &HFF00&) \ 2 ^ 8)
        gcTable(idx + 2) = CByte((v And &HFF0000) \ 2 ^ 16)
        idx = idx + 3
    Next
End Sub

Private Sub MakeGraphicControlExtension( _
    ByVal transparent As Boolean, _
    ByVal transparentIndex As Long, _
    ByRef gcExt As GraphicControlExtension)

    Dim transparentColorFlag As Long
    transparentColorFlag = IIf(transparent, 1, 0)

    With gcExt
        .ExtensionIntroducer = &H21
        .GraphicControlLabel = &HF9
        .BlockSize = &H4
        .PackedFields = transparentColorFlag
        .DelayTime = 0
        .TransparentColorIndex = transparentIndex
        .BlockTerminator = 0
    End With
End Sub

Private Sub MakeImageDescriptor( _
    ByVal pictWidth As Long, _
    ByVal pictHeight As Long, _
    ByVal bpp As Long, _
    ByRef imgDesc As ImageDescriptor)

    With imgDesc
        .ImageSeparator = &H2C
        .ImageLeftPosition = 0
        .ImageTopPosition = 0
        .ImageWidth = pictWidth
        .ImageHeight = pictHeight
        .PackedFields = 0
        .LZWMinimumCodeSize = IIf(bpp < 2, 2, bpp)
    End With
End Sub

Private Sub MakeImageBlocks( _
    ByRef data() As Byte, _
    ByVal bpp As Long, _
    ByRef imgBlocks() As ImageBlock)

    Dim compressedData() As Byte
    Call Compress(data, bpp, compressedData)

    Dim numBytes As Long
    numBytes = UBound(compressedData) + 1

    Dim quotient As Long
    quotient = numBytes \ &HFE&

    Dim remainder As Long
    remainder = numBytes Mod &HFE&

    ReDim imgBlocks(numBytes \ &HFE& + 1)

    Dim i As Long
    For i = 0 To quotient - 1
        With imgBlocks(i)
            .Size = &HFE&
            ReDim .BlockData(&HFE& - 1)
            Call MoveMemory(VarPtr(.BlockData(0)), VarPtr(compressedData(&HFE& * i)), &HFE&)
        End With
    Next

    If remainder > 0 Then
        With imgBlocks(quotient)
            .Size = remainder
            ReDim .BlockData(remainder - 1)
            Call MoveMemory(VarPtr(.BlockData(0)), VarPtr(compressedData(&HFE& * quotient)), remainder)
        End With
    End If

    imgBlocks(UBound(imgBlocks)).Size = 0 ' block terminator
End Sub

Private Sub InitializeDictionary(ByVal bpp As Long)
    Set m_dict = CreateObject("Scripting.Dictionary")
    m_bitsLen = 3

    If bpp < 2 Then
        bpp = 2
    End If

    Dim code As Long
    For code = 0 To (2 ^ bpp - 1)
        Call m_dict.Add("&H" & Hex(code), code)

        If code >= (2 ^ m_bitsLen - 1) Then
            m_bitsLen = m_bitsLen + 1
        End If
    Next

    m_clearCode = "&H" & Hex(m_dict.Count)
    code = m_dict.Count
    Call m_dict.Add(m_clearCode, code)

    m_endCode = "&H" & Hex(m_dict.Count)
    code = m_dict.Count
    Call m_dict.Add(m_endCode, code)
End Sub

Private Sub Compress(ByRef data() As Byte, ByVal bpp As Long, ByRef buffer() As Byte)
    Set m_buf = New List

    Call InitializeDictionary(bpp)
    Call m_buf.Add(Array(CLng(m_clearCode), m_bitsLen))

    Dim pfx As String
    Dim sfx As String
    Dim w   As String

    Dim i As Long
    i = 0

    pfx = "&H" & Hex(data(i))

    Dim maxIndex As Long
    maxIndex = UBound(data)

    Do While i < maxIndex
        i = i + 1
        sfx = "&H" & Hex(data(i))
        w = pfx & sfx

        If Not m_dict.Exists(w) Then
            Call m_dict.Add(w, m_dict.Count)
            Call m_buf.Add(Array(m_dict(pfx), m_bitsLen))
            pfx = sfx
            Call UpdateBitsLength(bpp)

            GoTo Continue
        End If

        Do While True
            i = i + 1
            If i > maxIndex Then Exit Do

            sfx = "&H" & Hex(data(i))

            If m_dict.Exists(w & sfx) Then
                w = w & sfx
            Else
                Call m_dict.Add(w & sfx, m_dict.Count)
                Call m_buf.Add(Array(m_dict(w), m_bitsLen))
                pfx = sfx
                Call UpdateBitsLength(bpp)

                Exit Do
            End If
        Loop
Continue:
    Loop

    Select Case i
        Case Is = maxIndex
            Call m_buf.Add(Array(m_dict("&H" & Hex(data(UBound(data)))), m_bitsLen))
        Case Is > maxIndex
            Call m_buf.Add(Array(m_dict(w), m_bitsLen))
        Case Else
            Call Err.Raise(51)
    End Select

    Call m_buf.Add(Array(CLng(m_endCode), m_bitsLen))

    Dim entries() As Variant
    entries = m_buf.Items()

    Dim bs As New BitSequence
    Call bs.Init(PackingOrder.LSBFirst)

    Dim j As Long
    For j = 0 To UBound(entries)
        Call bs.Append(entries(j)(0), entries(j)(1))
    Next

    buffer = bs.GetBytes()
End Sub

Private Sub UpdateBitsLength(ByVal bpp As Long)
    If (m_dict.Count - 1) > (2 ^ m_bitsLen - 1) Then
        m_bitsLen = m_bitsLen + 1
    End If

    If m_bitsLen > 12 Then
        Call m_buf.Add(Array(CLng(m_clearCode), 12))
        Call InitializeDictionary(bpp)
    End If
End Sub

Private Sub ToBytes( _
    ByRef hdr As GifHeader, _
    ByRef lsDesc As LogicalScreenDescriptor, _
    ByRef gcTable() As Byte, _
    ByRef gcExt As GraphicControlExtension, _
    ByRef imgDesc As ImageDescriptor, _
    ByRef imgBlocks() As ImageBlock, _
    ByRef trailer As Byte, _
    ByRef buffer() As Byte)

    Dim hdrSize As Long
    hdrSize = 6

    Dim lsDescSize As Long
    lsDescSize = 7

    Dim gcTableSize As Long
    gcTableSize = UBound(gcTable) + 1

    Dim gcExtSize As Long
    If (gcExt.PackedFields And 1) > 0 Then
        gcExtSize = 8
    Else
        gcExtSize = 0
    End If

    Dim imgDescSize As Long
    imgDescSize = 11

    Dim imgBlocksSize As Long
    Dim i As Long
    For i = 0 To UBound(imgBlocks)
        imgBlocksSize = imgBlocksSize + 1
        If imgBlocks(i).Size > 0 Then
            imgBlocksSize = imgBlocksSize + (UBound(imgBlocks(i).BlockData) + 1)
        End If
    Next

    Dim trSize As Long
    trSize = 1

    Dim sz As Long
    sz = hdrSize + lsDescSize + gcTableSize + gcExtSize + imgDescSize + imgBlocksSize + trSize

    ReDim buffer(sz - 1)

    Dim idx As Long
    idx = 0

    With hdr
        Call MoveMemory(VarPtr(buffer(idx)), VarPtr(.ExtensionIntroducer(0)), 3)
        idx = idx + 3
        Call MoveMemory(VarPtr(buffer(idx)), VarPtr(.Version(0)), 3)
        idx = idx + 3
    End With

    Call MoveMemory(VarPtr(buffer(idx)), VarPtr(lsDesc), lsDescSize)
    idx = idx + lsDescSize

    Call MoveMemory(VarPtr(buffer(idx)), VarPtr(gcTable(0)), gcTableSize)
    idx = idx + gcTableSize

    If (gcExt.PackedFields And 1) > 0 Then
        Call MoveMemory(VarPtr(buffer(idx)), VarPtr(gcExt.ExtensionIntroducer), gcExtSize)
        idx = idx + gcExtSize
    End If

    With imgDesc
        Call MoveMemory(VarPtr(buffer(idx)), VarPtr(.ImageSeparator), 1)
        idx = idx + 1
        Call MoveMemory(VarPtr(buffer(idx)), VarPtr(.ImageLeftPosition), 2)
        idx = idx + 2
        Call MoveMemory(VarPtr(buffer(idx)), VarPtr(.ImageTopPosition), 2)
        idx = idx + 2
        Call MoveMemory(VarPtr(buffer(idx)), VarPtr(.ImageWidth), 2)
        idx = idx + 2
        Call MoveMemory(VarPtr(buffer(idx)), VarPtr(.ImageHeight), 2)
        idx = idx + 2
        Call MoveMemory(VarPtr(buffer(idx)), VarPtr(.PackedFields), 1)
        idx = idx + 1
        Call MoveMemory(VarPtr(buffer(idx)), VarPtr(.LZWMinimumCodeSize), 1)
        idx = idx + 1
    End With

    For i = 0 To UBound(imgBlocks)
        With imgBlocks(i)
            Call MoveMemory(VarPtr(buffer(idx)), VarPtr(.Size), 1)
            idx = idx + 1
            If .Size > 0 Then
                Call MoveMemory(VarPtr(buffer(idx)), VarPtr(.BlockData(0)), UBound(.BlockData) + 1)
                idx = idx + UBound(.BlockData) + 1
            End If
        End With
    Next

    Call MoveMemory(VarPtr(buffer(idx)), VarPtr(trailer), trSize)
End Sub
