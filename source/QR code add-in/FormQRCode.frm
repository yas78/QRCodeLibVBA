VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormQRCode 
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2070
   OleObjectBlob   =   "FormQRCode.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "FormQRCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const DEFAULT_WIDTH  As Single = 162
Const DEFAULT_HEIGHT As Single = 162
Const EXP_SIZE       As Single = 80

Private m_minWidth   As Single
Private m_minHeight  As Single
Private m_maxWidth   As Single
Private m_maxHeight  As Single

Public Sub ShowForm(ByVal pic As IPictureDisp, Optional ByVal s As String = "")
    Me.Caption = s
    Set Me.Picture = pic
    Call Me.Show(vbModeless)
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = KeyCodeConstants.vbKeyEscape And Shift = 0 Then
        Unload Me
    End If
End Sub

Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If m_maxWidth >= Me.Width + EXP_SIZE Then
        Me.Width = Me.Width + EXP_SIZE
        Me.Height = Me.Height + EXP_SIZE
        
        Me.Top = Me.Top - (EXP_SIZE / 2)
        Me.Left = Me.Left - (EXP_SIZE / 2)
    End If
End Sub

Private Sub UserForm_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Select Case Button
        Case 1
            If m_maxWidth >= Me.Width + EXP_SIZE Then
                Me.Width = Me.Width + EXP_SIZE
                Me.Height = Me.Height + EXP_SIZE
                Me.Top = Me.Top - (EXP_SIZE / 2)
                Me.Left = Me.Left - (EXP_SIZE / 2)
            End If
        Case 2
            If m_minWidth <= Me.Width - EXP_SIZE Then
                Me.Width = Me.Width - EXP_SIZE
                Me.Height = Me.Height - EXP_SIZE
                
                Me.Top = Me.Top + (EXP_SIZE / 2)
                Me.Left = Me.Left + (EXP_SIZE / 2)
            End If
    End Select
End Sub

Private Sub UserForm_Initialize()
    Dim merginWidth  As Single
    Dim merginHeight As Single
    
    merginWidth = Me.Width - Me.InsideWidth
    merginHeight = Me.Height - Me.InsideHeight
    
    m_minWidth = merginWidth + DEFAULT_WIDTH
    m_minHeight = merginHeight + DEFAULT_HEIGHT
    m_maxWidth = merginWidth + (DEFAULT_WIDTH * 4)
    m_maxHeight = merginHeight + (DEFAULT_HEIGHT * 4)
    
    Me.Width = merginWidth + DEFAULT_WIDTH
    Me.Height = merginHeight + DEFAULT_HEIGHT
End Sub

Private Sub UserForm_Deactivate()
    Unload Me
End Sub

