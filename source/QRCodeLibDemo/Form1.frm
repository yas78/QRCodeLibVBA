VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form1 
   Caption         =   "QR Code"
   ClientHeight    =   8910
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10590
   OleObjectBlob   =   "Form1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const DEFAULT_MODULE_SIZE As Long = 5
Private Const DEFAULT_VERSION     As Long = 40
Private Const IMAGE_WIDTH         As Long = 166
Private Const IMAGE_HEIGHT        As Long = 166
Private Const IMAGE_MARGIN        As Long = 2
Private Const COL_COUNT           As Long = 3

Private Sub Update_fraQRCodeImage()
    btnSave.Enabled = False
    Call fraQRCodeImage.Controls.Clear
    fraQRCodeImage.ScrollHeight = 0

    If txtData.TextLength = 0 Then Exit Sub

    Dim ecLevel As ErrorCorrectionLevel
    ecLevel = cmbErrorCorrectionLevel.Value
    
    Dim sz As Long
    sz = CLng(txtModuleSize.Text)
    
    Dim foreRGB As String
    foreRGB = "#" & txtForeColor.Text
    
    Dim backRGB As String
    backRGB = "#" & txtBackColor.Text
    
    Dim maxVer As Long
    maxVer = CLng(cmbMaxVersion.Text)
    
    Dim structAppend As Boolean
    structAppend = chkStructuredAppend.Value
    
    Dim encMode As String
    encMode = cmbEncoding.Value

On Error GoTo Catch
    Dim sbls As QRCodeLib.Symbols
    Set sbls = CreateSymbols(ecLevel, maxVer, structAppend, encMode)
    Call sbls.AppendText(txtData.Text)

    Dim sbl As Symbol
    Dim ctl As Control
    Dim img As Image
    Dim idx As Long

    For idx = 0 To sbls.Count - 1
        Set sbl = sbls(idx)
        Set ctl = Me.fraQRCodeImage.Controls.Add("Forms.Image.1")
        With ctl
            .Left = (IMAGE_WIDTH + IMAGE_MARGIN) * (idx Mod COL_COUNT) + IMAGE_MARGIN
            .Top = (IMAGE_WIDTH + IMAGE_MARGIN) * (idx \ COL_COUNT) + IMAGE_MARGIN
            .Width = IMAGE_WIDTH
            .Height = IMAGE_HEIGHT
        End With

        Set img = ctl
        img.PictureSizeMode = fmPictureSizeModeStretch
        img.BorderStyle = fmBorderStyleNone
        img.Picture = sbl.GetPicture(sz, True, foreRGB, backRGB)
    Next

    fraQRCodeImage.ScrollHeight = _
        CLng((sbls.Count + 3) \ COL_COUNT) * (IMAGE_HEIGHT + IMAGE_MARGIN) + IMAGE_MARGIN
    btnSave.Enabled = txtData.TextLength > 0

Finally:
On Error GoTo 0
    Exit Sub

Catch:
    Call MsgBox(Err.Description, vbExclamation, "")
    Resume Finally
End Sub

Private Sub Set_txtModuleSize(ByVal moduleSize As Long)
    txtModuleSize.Text = CStr(moduleSize)
    spbModuleSize.Value = moduleSize
End Sub

Private Sub btnSave_Click()
    Dim ecLevel As ErrorCorrectionLevel
    ecLevel = cmbErrorCorrectionLevel.Value
    
    Dim sz As Long
    sz = CLng(txtModuleSize.Text)
    
    Dim foreRGB As String
    foreRGB = "#" & txtForeColor.Text
    
    Dim backRGB As String
    backRGB = "#" & txtBackColor.Text
    
    Dim maxVer As Long
    maxVer = CLng(cmbMaxVersion.Text)
    
    Dim structAppend As Boolean
    structAppend = chkStructuredAppend.Value
    
    Dim encMode As String
    encMode = cmbEncoding.Value

    Dim fs As New FileSystemObject

    Dim fileFilters As String
    fileFilters = "Monochrome Bitmap (*.bmp), *.bmp, SVG (*.svg), *svg"

    Dim fBaseName As Variant
    fBaseName = Application.GetSaveAsFilename("", fileFilters)

    If VarType(fBaseName) = vbBoolean Then Exit Sub
    
    Dim ext As String
    ext = "." & fs.GetExtensionName(fBaseName)
    
    fBaseName = fs.GetParentFolderName(fBaseName) & "\" & fs.GetBaseName(fBaseName)

On Error GoTo Catch
    Dim sbls As Symbols
    Set sbls = CreateSymbols(ecLevel, maxVer, structAppend, encMode)
    Call sbls.AppendText(txtData.Text)

    Dim filePath As String
    Dim sbl      As Symbol
    Dim i        As Long

    For i = 0 To sbls.Count - 1
        Set sbl = sbls(i)

        If sbls.Count = 1 Then
            filePath = fBaseName & ext
        Else
            filePath = fBaseName & "_" & CStr(i + 1) & ext
        End If

        If fs.FileExists(filePath) Then
            Call fs.DeleteFile(filePath)
        End If
        
        Select Case LCase(ext)
            Case ".bmp"
                Call sbl.SaveBitmap(filePath, sz, True, foreRGB, backRGB)
            Case ".svg"
                Call sbl.SaveSvg(filePath, sz, foreRGB)
            Case Else
                Call Err.Raise(51)
        End Select
    Next

Finally:
On Error GoTo 0
    Exit Sub

Catch:
    Call MsgBox(Err.Description, vbExclamation, "")
    Resume Finally
End Sub

Private Sub chkStructuredAppend_Change()
    Call Update_fraQRCodeImage
End Sub

Private Sub cmbEncoding_Change()
    Call Update_fraQRCodeImage
End Sub

Private Sub cmbErrorCorrectionLevel_Change()
    Call Update_fraQRCodeImage
End Sub

Private Sub cmbMaxVersion_Change()
    Call Update_fraQRCodeImage
End Sub

Private Sub spbModuleSize_Change()
    Call Set_txtModuleSize(spbModuleSize.Value)
End Sub

Private Sub txtForeColor_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    txtForeColor.Text = Left$(txtForeColor.Text & String(6, "0"), 6)

    If txtForeColor.Text Like "*[!0-9A-Fa-f]*" Then
        txtForeColor.Text = "000000"
    End If
End Sub

Private Sub txtBackColor_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    txtBackColor.Text = Left$(txtBackColor.Text & String(6, "0"), 6)

    If txtBackColor.Text Like "*[!0-9A-Fa-f]*" Then
        txtBackColor.Text = "FFFFFF"
    End If
End Sub

Private Sub txtModuleSize_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    If txtModuleSize.TextLength = 0 Then
        Call Set_txtModuleSize(DEFAULT_MODULE_SIZE)
    End If

    If txtModuleSize.Text Like "*[!0-9]*" Then
        Call Set_txtModuleSize(DEFAULT_MODULE_SIZE)
    End If

    If CLng(txtModuleSize.Text) < 2 Then
        Call Set_txtModuleSize(2)
    End If

    If CLng(txtModuleSize.Text) > 100 Then
        Call Set_txtModuleSize(100)
    End If

    spbModuleSize.Value = CLng(txtModuleSize.Text)
End Sub

Private Sub txtData_Change()
    Call Update_fraQRCodeImage
End Sub

Private Sub txtModuleSize_KeyDown( _
    ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    If txtModuleSize.TextLength = 0 Then Exit Sub
    If txtModuleSize.Text Like "*[!0-9]*" Then Exit Sub

    Dim sz As Long
    sz = CLng(txtModuleSize.Text)

    Select Case KeyCode
        Case 38
            If 2 <= sz And sz < 20 Then
                txtModuleSize.Text = CStr(sz + 1)
            End If
        Case 40
            If 2 < sz And sz <= 20 Then
                txtModuleSize.Text = CStr(sz - 1)
            End If
    End Select
End Sub

Private Sub UserForm_Initialize()
    With cmbErrorCorrectionLevel
        .ColumnCount = 2
        .ColumnWidths = "0"
        .TextColumn = 2
        .BoundColumn = 1

        .AddItem
        .List(0, 0) = ErrorCorrectionLevel.L
        .List(0, 1) = "L (7%)"

        .AddItem
        .List(1, 0) = ErrorCorrectionLevel.M
        .List(1, 1) = "M (15%)"

        .AddItem
        .List(2, 0) = ErrorCorrectionLevel.Q
        .List(2, 1) = "Q (25%)"

        .AddItem
        .List(3, 0) = ErrorCorrectionLevel.H
        .List(3, 1) = "H (30%)"

        .ListIndex = 1
    End With

    Call cmbEncoding.AddItem("Shift_JIS")
    Call cmbEncoding.AddItem("UTF-8")
    cmbEncoding.ListIndex = 0

    Dim i As Long
    For i = 1 To 40
        Call cmbMaxVersion.AddItem(i)
    Next

    cmbMaxVersion.Value = DEFAULT_VERSION

    Call Set_txtModuleSize(DEFAULT_MODULE_SIZE)
    chkStructuredAppend.Value = False

    txtForeColor.Text = "000000"
    txtBackColor.Text = "FFFFFF"

    btnSave.Enabled = False
End Sub

