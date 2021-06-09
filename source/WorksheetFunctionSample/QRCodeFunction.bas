Attribute VB_Name = "QRCodeFunction"
Option Explicit

Private Const TemporaryFolder = 2
Private m_fs As Object

Public Function QR(ByVal s As String, _
                   Optional ByVal charsetNmae As String = "Shift_JIS") As Variant
    If m_fs Is Nothing Then
        Set m_fs = CreateObject("Scripting.FileSystemObject")
    End If

On Error GoTo Catch
    If Not (TypeOf Application.Caller Is Range) Then Exit Function

    Dim rng As Range
    Set rng = Application.Caller.MergeArea
    Call DeleteShape(rng)

    If Len(s) = 0 Then Exit Function

    Dim sbls As Symbols
    Set sbls = CreateSymbols(charsetName:=charsetNmae)
    Call sbls.AppendText(s)

    Dim filePath As String
    filePath = m_fs.GetSpecialFolder(TemporaryFolder) & "\" & m_fs.GetTempName()

    If m_fs.FileExists(filePath) Then Call m_fs.DeleteFile(filePath)
    Call sbls(0).SaveAs(filePath)
    Call AddPicture(filePath, rng)

    QR = ""

Finally:
On Error GoTo 0
    If m_fs.FileExists(filePath) Then Call m_fs.DeleteFile(filePath)
    Exit Function
Catch:
    QR = CVErr(xlErrValue)
    Resume Finally
End Function

Private Function DeleteShape(ByVal Targe As Range)
    Dim ws As Worksheet
    Set ws = Targe.Parent

    Dim shps As Shapes
    Set shps = ws.Shapes

    Dim shp As Shape
    Dim rng As Range
    Dim i As Long
    For i = shps.Count To 1 Step -1
        Set shp = shps(i)
        Set rng = ws.Range(shp.TopLeftCell, shp.BottomRightCell)
        If Not (Intersect(Targe, rng) Is Nothing) Then shp.Delete
    Next
End Function

Private Sub AddPicture(ByVal filePath As String, ByVal Target As Range)
    Dim shps As Shapes
    Set shps = Target.Parent.Shapes

    Dim shp As Shape
    Set shp = shps.AddPicture(filePath, msoFalse, msoTrue, Target.Left, Target.Top, 0, 0)

    With shp
        Call .ScaleHeight(1, msoTrue)
        Call .ScaleWidth(1, msoTrue)
        .LockAspectRatio = msoTrue
        .Placement = xlMoveAndSize
    End With

    If Target.Height * shp.Width <= Target.Width * shp.Height Then
        shp.Height = Target.Height - 10
    Else
        shp.Width = Target.Width - 10
    End If

    shp.Left = Target.Left + Target.Width / 2 - shp.Width / 2
    shp.Top = Target.Top + Target.Height / 2 - shp.Height / 2
End Sub
