VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit


Private Const ID As String = "QRCodeAddIn"

Private Sub Workbook_AddinInstall()
    Call CreateMenu
End Sub

Private Sub Workbook_AddinUninstall()
    Call DeleteMenu
End Sub

Private Sub Workbook_Open()
    Call CreateMenu
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Call DeleteMenu
End Sub

Private Sub CreateMenu()
    If MenuExists() Then Exit Sub

    Dim cmdBar      As CommandBar
    Dim cmdBarCtrl  As CommandBarControl
    Dim menuBtn     As CommandBarButton

    For Each cmdBar In Application.CommandBars
        If cmdBar.Name = "Cell" Then
            Set menuBtn = cmdBar.Controls.Add(Type:=msoControlButton)
            menuBtn.Tag = ID
            menuBtn.Caption = "QR code(&Q)"
            menuBtn.OnAction = "ShowQRCodeWindow"
            menuBtn.Style = msoButtonCaption
        End If
    Next
End Sub

Private Sub DeleteMenu()
    Dim cmdBarCtrl  As CommandBarControl
    Dim cmdBar      As CommandBar

    For Each cmdBar In Application.CommandBars
        If cmdBar.Name = "Cell" Then
            For Each cmdBarCtrl In cmdBar.Controls
                If cmdBarCtrl.Tag = ID Then
                    Call cmdBarCtrl.Delete
                End If
            Next
        End If
    Next
End Sub

Private Function MenuExists() As Boolean
    Dim cmdBar      As CommandBar
    Dim cmdBarCtrl  As CommandBarControl

    MenuExists = False

    For Each cmdBar In Application.CommandBars
        If cmdBar.Name = "Cell" Then
            For Each cmdBarCtrl In cmdBar.Controls
                If cmdBarCtrl.Tag = ID Then
                    MenuExists = True
                    Exit Function
                End If
            Next
        End If
    Next
End Function

