Attribute VB_Name = "UserFunction"
Option Explicit


Private Const DEFAULT_MODULE_SIZE As Long = 4

Public Function QRcode(ByVal Text As String, _
                       Optional ByVal ModuleSize As Long = DEFAULT_MODULE_SIZE, _
                       Optional ByVal ForeColor As String = "#000000", _
                       Optional ByVal BackColor As String = "#FFFFFF", _
                       Optional ByVal VersionLimit As Long = MAX_VERSION) As Variant
On Error GoTo Catch
    If TypeName(Application.Caller) <> "Range" Then Exit Function
    If Len(Text) = 0 Then Call Err.Raise(5)
    If Not (1 <= ModuleSize And ModuleSize <= 20) Then Call Err.Raise(5)
    If Not (ColorCode.Valid(ForeColor)) Then Err.Raise (5)
    If Not (ColorCode.Valid(BackColor)) Then Err.Raise (5)
    If Not (MIN_VERSION <= VersionLimit And VersionLimit <= MAX_VERSION) Then Call Err.Raise(5)
    
    Dim filePath As String
    filePath = Directory.GetTempPath()
    
    If Directory.FileExists(filePath) Then
        Call Directory.DeleteFile(filePath)
    End If
            
    Dim rng As Range
    Set rng = Application.Caller.MergeArea
    
    Call DeleteShape(rng)
    
    Dim sbls As Symbols
    Set sbls = CreateSymbols(maxVer:=VersionLimit)
    Call sbls.AppendText(Text)
    
    Call sbls(0).SaveToFile(filePath, ModuleSize, ForeColor, BackColor, True)
    
    Dim shp As Shape
    Set shp = AddPicture(filePath, rng)
    Call FitToCell(shp, rng)
    
    Dim ret As String
    ret = "QR code"
    
    QRcode = ret

Finally:
On Error GoTo 0
    If Directory.FileExists(filePath) Then
        Call Directory.DeleteFile(filePath)
    End If

    Exit Function

Catch:
    QRcode = CVErr(xlErrValue)
    Resume Finally
End Function

Private Sub DeleteShape(ByVal rng As Range)
    Dim shp As Shape
    
    For Each shp In rng.Parent.Shapes
        If rng.Left <= shp.Left And (shp.Left + shp.Width) < (rng.Left + rng.Width) And _
            rng.Top <= shp.Top And (shp.Top + shp.Height) < (rng.Top + rng.Height) Then
            Call shp.Delete
        End If
    Next
End Sub

Private Function AddPicture(ByVal filePath As String, ByVal rng As Range) As Shape
    Dim shp As Shape
    Set shp = rng.Parent.Shapes.AddPicture(filePath, msoFalse, msoTrue, rng.Left, rng.Top, 0, 0)
    
    Call shp.ScaleHeight(1, msoTrue)
    Call shp.ScaleWidth(1, msoTrue)
    shp.LockAspectRatio = msoTrue
    shp.Placement = IIf(Application.Version = EXCEL_2010, xlMoveAndSize, xlMove)

    Set AddPicture = shp
End Function

Private Sub FitToCell(ByVal shp As Shape, ByVal rng As Range)
    If rng.Height * shp.Width <= rng.Width * shp.Height Then
        shp.Height = rng.Height - 4
    Else
        shp.Width = rng.Width - 4
    End If

    shp.Left = rng.Left + rng.Width / 2 - shp.Width / 2
    shp.Top = rng.Top + rng.Height / 2 - shp.Height / 2
End Sub
