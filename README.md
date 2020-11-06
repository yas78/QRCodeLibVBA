# QRCodeLibVBA
QRCodeLibVBAは、Excel VBAで書かれたQRコード生成ライブラリです。  
JIS X 0510に基づくモデル２コードシンボルを生成します。

## 特徴
- 数字・英数字・8ビットバイト・漢字モードに対応しています
- 分割QRコードを作成可能です
- 1bppまたは24bpp BMPファイル(DIB)へ保存可能です
- SVG形式で保存可能です
- 1bppまたは24bpp IPictureDispオブジェクトとして取得可能です  
- 画像の配色(前景色・背景色)を指定可能です
- 8ビットバイトモードでの文字コードを指定可能です
- QRコード画像をクリップボードに保存可能です
- Excel 32bit, 64bitの両環境で使用可能です


## クイックスタート
QRCodeLib.xlam を参照設定してください。  


## 使用方法
### 例１．単一シンボルで構成される(分割QRコードではない)QRコードの、最小限のコードを示します。

```VBA
Public Sub Example1()
    Dim sbls As Symbols
    Set sbls = CreateSymbols()
    sbls.AppendText "012345abcdefg"

    Dim pict As stdole.IPictureDisp
    Set pict = sbls(0).GetPicture()
End Sub
```

### 例２．誤り訂正レベルを指定する
CreateSymbols関数の引数に、ErrorCorrectionLevel列挙型の値を設定してSymbolsオブジェクトを生成します。

```VBA
Public Sub Example2()
    Dim sbls As Symbols
    Set sbls = CreateSymbols(ErrorCorrectionLevel.H)
End Sub
```

### 例３．型番の上限を指定する
CreateSymbols関数の maxVer 引数を設定してSymbolsオブジェクトを生成します。

```VBA
Public Sub Example3()
    Dim sbls As Symbols
    Set sbls = CreateSymbols(maxVer:=10)
End Sub
```

### 例４．8ビットバイトモードで使用する文字コードを指定する
CreateSymbols関数の byteModeCharsetName 引数を設定してSymbolsオブジェクトを生成します。

```VBA
Public Sub Example4()
    Dim sbls As Symbols
    Set sbls = CreateSymbols(byteModeCharsetName:="utf-8")
End Sub
```

### 例５．分割QRコードを作成する
CreateSymbols関数の引数を設定してSymbolsオブジェクトを生成します。型番の上限を指定しない場合は、型番40を上限として分割されます。  

型番1を超える場合に分割し、各QRコードのIPictureDispオブジェクトを取得する例を示します。

```VBA
Public Sub Example5()
    Dim sbls As Symbols
    Set sbls = CreateSymbols(maxVer:=1, allowStructuredAppend:=True)
    sbls.AppendText "abcdefghijklmnopqrstuvwxyz"
    
    Dim pict As stdole.IPictureDisp
    Dim sbl As Symbol
    
    For Each sbl In sbls
        Set pict = sbl.GetPicture()
    Next
End Sub
```

### 例６．BMPファイルへ保存する
SymbolクラスのSaveBitmapメソッドを使用します。

```VBA
Public Sub Example6()
    Dim sbls As Symbols
    Set sbls = CreateSymbols()
    sbls.AppendText "012345abcdefg"
    
    ' 24bpp DIB
    sbls(0).SaveBitmap "D:\QRcode.bmp"
    
    ' 10 pixels per module
    sbls(0).SaveBitmap "D:\QRcode.bmp", moduleSize:=10
    
    ' Specify foreground and background colors.
    sbls(0).SaveBitmap "D:\QRcode.bmp", foreRGB:="#0000FF", backRGB:="#FFFF00"
    
    ' 1bpp DIB
    sbls(0).SaveBitmap "D:\QRcode.bmp", monochrome:=True
End Sub
```

### 例７．SVGファイルへ保存する
SymbolクラスのSaveSvgメソッドを使用します。

```VBA
Public Sub Example6()
    Dim sbls As Symbols
    Set sbls = CreateSymbols()
    sbls.AppendText "012345abcdefg"
    
    sbls(0).SaveBitmap "D:\QRcode.svg"    
End Sub
```

### 例８．SVGデータを取得する
SymbolクラスのSaveSvgメソッドを使用します。

```VBA
Public Sub Example6()
    Dim sbls As Symbols
    Set sbls = CreateSymbols()
    sbls.AppendText "012345abcdefg"
    
    sbls(0).GetSvg "D:\QRcode.svg"    
End Sub
```

### 例９．クリップボードへ格納する
SymbolクラスのSetToClipboardメソッドを使用します。

```VBA
Public Sub Example7()
    Dim sbls As Symbols
    Set sbls = CreateSymbols()
    sbls.AppendText "012345abcdefg"
    
    sbls(0).SetToClipBoard
    sbls(0).SetToClipBoard moduleSize:=10
    sbls(0).SetToClipBoard foreRGB:="#0000FF", backRGB:="#FFFF00"
End Sub
```

