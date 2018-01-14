VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form1 
   Caption         =   "QR Code"
   ClientHeight    =   8430
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
Private Const IMAGE_WIDTH         As Long = 122
Private Const IMAGE_HEIGHT        As Long = 122
Private Const IMAGE_MARGIN        As Long = 2

Private Sub Update_fraQRCodeImage()
    
    btnSave.Enabled = False
    Call fraQRCodeImage.Controls.Clear
    fraQRCodeImage.ScrollHeight = 0
    
    If txtData.TextLength = 0 Then Exit Sub
    
    Dim ecLevel As ErrorCorrectionLevel
    ecLevel = cmbErrorCorrectionLevel.Value

On Error GoTo Catch_
    Dim sbls As QRCodeLib.Symbols
    Set sbls = CreateSymbols(ecLevel, _
                             CLng(cmbMaxVersion.Text), _
                             chkStructuredAppend.Value, _
                             cmbEncoding.Value)
    Call sbls.AppendString(txtData.Text)
    
    Dim sbl As Symbol
    Dim ctl As Control
    Dim img As Image
    Dim idx As Long
    
    For idx = 0 To sbls.Count - 1
        Set sbl = sbls.Item(idx)
        Set ctl = Me.fraQRCodeImage.Controls.Add("Forms.Image.1")
        
        ctl.Left = (IMAGE_WIDTH + IMAGE_MARGIN) * (idx Mod 4) + IMAGE_MARGIN
        ctl.Top = (IMAGE_WIDTH + IMAGE_MARGIN) * (idx \ 4) + IMAGE_MARGIN
        ctl.Width = IMAGE_WIDTH
        ctl.Height = IMAGE_HEIGHT
        
        Set img = ctl
        img.PictureSizeMode = fmPictureSizeModeStretch
        img.BorderStyle = fmBorderStyleNone
        img.Picture = sbl.Get24bppImage(CLng(txtModuleSize.Text))
    Next
    
    fraQRCodeImage.ScrollHeight = _
        CLng((sbls.Count + 3) \ 4) * (IMAGE_HEIGHT + IMAGE_MARGIN) + _
        IMAGE_MARGIN
    btnSave.Enabled = txtData.TextLength > 0

Finally_:
On Error GoTo 0
    Exit Sub

Catch_:
    Call MsgBox(Err.Description, vbExclamation, "")
    Resume Finally_

End Sub

Private Sub Set_txtModuleSize(ByVal moduleSize As Long)
    
    txtModuleSize.Text = CStr(moduleSize)
    spbModuleSize.Value = moduleSize

End Sub

Private Sub btnSave_Click()

    Dim fs As New FileSystemObject
    Dim fBaseName As Variant
    fBaseName = Application.GetSaveAsFilename("", "Monochrome Bitmap, *.bmp")
    
    If VarType(fBaseName) = vbBoolean Then Exit Sub
    
    fBaseName = fs.GetParentFolderName(fBaseName) & "\" & fs.GetBaseName(fBaseName)
    
    Dim ecLevel As ErrorCorrectionLevel
    ecLevel = cmbErrorCorrectionLevel.Value
    
On Error GoTo Catch_
    
    Dim sbls As Symbols
    Set sbls = CreateSymbols(ecLevel, _
                             CLng(cmbMaxVersion.Text), _
                             chkStructuredAppend.Value, _
                             cmbEncoding.Value)
    Call sbls.AppendString(txtData.Text)
    
    Dim filePath As String
    Dim sbl      As Symbol
    Dim i        As Long
    
    For i = 0 To sbls.Count - 1
        Set sbl = sbls.Item(i)
        
        If sbls.Count = 1 Then
            filePath = fBaseName & ".bmp"
        Else
            filePath = fBaseName & "_" & CStr(i + 1) & ".bmp"
        End If
        
        If fs.FileExists(filePath) Then
            Call fs.DeleteFile(filePath)
        End If
        
        Call sbl.Save1bppDIB(filePath, CLng(txtModuleSize.Text))
    Next

Finally_:
On Error GoTo 0
    Exit Sub
    
Catch_:
    Call MsgBox(Err.Description, vbExclamation, "")
    Resume Finally_
    
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

Private Sub txtModuleSize_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

    If txtModuleSize.TextLength = 0 Then
        Call Set_txtModuleSize(DEFAULT_MODULE_SIZE)
    End If
    
    If txtModuleSize.Text Like "*[!0-9]*" Then
        Call Set_txtModuleSize(DEFAULT_MODULE_SIZE)
    End If
    
    If CLng(txtModuleSize.Text) = 0 Then
        Call Set_txtModuleSize(1)
    End If
                
    If CLng(txtModuleSize.Text) > 20 Then
        Call Set_txtModuleSize(20)
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
            If sz >= 1 And sz < 20 Then
                txtModuleSize.Text = CStr(sz + 1)
            End If
        Case 40
            If sz > 1 And sz <= 20 Then
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
    
    cmbEncoding.AddItem "Shift_JIS"
    cmbEncoding.AddItem "UTF-8"
    cmbEncoding.ListIndex = 0
    
    Dim i As Long
    
    For i = QRCodeLib.Constants.MIN_VERSION To QRCodeLib.Constants.MAX_VERSION
        cmbMaxVersion.AddItem i
    Next

    cmbMaxVersion.Value = DEFAULT_VERSION
    
    Call Set_txtModuleSize(DEFAULT_MODULE_SIZE)
    chkStructuredAppend.Value = False

    btnSave.Enabled = False
    
End Sub
