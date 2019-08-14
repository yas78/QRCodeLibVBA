Attribute VB_Name = "Main"
Option Explicit

Private Sub ShowQRCodeWindow()
On Error GoTo Catch
    Dim dataText As String
    dataText = ActiveCell.Text
    
    If Len(dataText) = 0 Then Exit Sub
    
    Dim sbls As Symbols
    Set sbls = CreateSymbols()
    Call sbls.AppendText(dataText)
    
    Dim sbl As Symbol
    Set sbl = sbls(0)

    Dim frm As FormQRCode
    Set frm = New FormQRCode
    Call frm.ShowForm(sbl.GetPicture())

Finally:
On Error GoTo 0
    Exit Sub

Catch:
    Call MsgBox(Err.Description, vbExclamation)
    Resume Finally
End Sub

