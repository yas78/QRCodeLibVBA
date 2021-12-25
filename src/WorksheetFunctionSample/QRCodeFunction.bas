Attribute VB_Name = "QRCodeFunction"
Option Explicit

Private Const TemporaryFolder = 2
Private m_fs As Object

Public Function QR(ByVal s As String, _
                   Optional ByVal charsetName As String = "Shift_JIS") As Variant
    If m_fs Is Nothing Then
        Set m_fs = CreateObject("Scripting.FileSystemObject")
    End If

On Error GoTo Catch
    If Not (TypeOf Application.Caller Is Range) Then Exit Function

    Dim rng As Range
    Set rng = Application.Caller.MergeArea
    Call DeletePictures(rng)

    If Len(s) = 0 Then Exit Function

    Dim sbls As Symbols
    Set sbls = CreateSymbols(charsetName:=charsetName)
    Call sbls.AppendText(s)

    Dim filePath As String
    filePath = m_fs.GetSpecialFolder(TemporaryFolder) & "\" & m_fs.GetTempName()

    If m_fs.FileExists(filePath) Then Call m_fs.DeleteFile(filePath)
    Call sbls(0).SaveAs(filePath, fmt:=fmtEMF)

    Dim shp As Shape
    Set shp = AddPicture(filePath, rng)
    Call FitToCell(shp)
    Call FillShape(shp, vbWhite)

    QR = ""

Finally:
On Error GoTo 0
    If m_fs.FileExists(filePath) Then Call m_fs.DeleteFile(filePath)
    Exit Function
Catch:
    QR = CVErr(xlErrValue)
    Resume Finally
End Function

Private Sub DeletePictures(ByVal Target As Range)
    Dim ws As Worksheet
    Set ws = Target.Parent

    Dim shps As Shapes
    Set shps = ws.Shapes

    Dim shp As Shape
    Dim rng As Range
    Dim i As Long
    For i = shps.Count To 1 Step -1
        Set shp = shps(i)
        If shp.Type = msoPicture Then
            Set rng = ws.Range(shp.TopLeftCell, shp.BottomRightCell)
            If Not (Intersect(Target, rng) Is Nothing) Then shp.Delete
        End If
    Next
End Sub

Private Function AddPicture(ByVal filePath As String, ByVal Target As Range) As Shape
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

    Set AddPicture = shp
End Function

Private Sub FillShape(ByVal shp As Shape, ByVal colorRgb As Long)
    With shp.Fill
        .ForeColor.RGB = colorRgb
        .Transparency = 0
        .Solid
    End With
End Sub

Private Sub FitToCell(ByVal shp As Shape)
    Dim rng As Range
    Set rng = shp.TopLeftCell.MergeArea

    If rng.Height * shp.Width <= rng.Width * shp.Height Then
        shp.Height = rng.Height - 10
    Else
        shp.Width = rng.Width - 10
    End If

    shp.Left = rng.Left + rng.Width / 2 - shp.Width / 2
    shp.Top = rng.Top + rng.Height / 2 - shp.Height / 2
End Sub
