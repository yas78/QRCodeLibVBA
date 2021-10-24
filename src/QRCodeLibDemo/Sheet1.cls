VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Const ADDR_DATA As String = "I4"
Const MODULE_SIZE As Long = 12

Private Sub btnForm_Click()
    Dim frm As Form1
    Set frm = New Form1
    Call frm.Show(vbModeless)
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    If Intersect(Target, Me.Range(ADDR_DATA)) Is Nothing Then Exit Sub

    Dim s As String
    s = Me.Range(ADDR_DATA).Value

    If Len(s) = 0 Then Exit Sub

On Error GoTo Catch
    Dim sbls As Symbols
    Set sbls = CreateSymbols()
    Call sbls.AppendText(s)

    With Me.imgQr
        .PictureAlignment = fmPictureAlignmentCenter
        .PictureSizeMode = fmPictureSizeModeStretch
        .Picture = sbls(0).GetPicture(MODULE_SIZE, pictype:=OlePicType.EnhMetaFile)
    End With

Finally:
On Error GoTo 0
    Exit Sub

Catch:
    MsgBox Err.Description, vbExclamation
    Resume Finally
End Sub
