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
Dim sbls As Symbols
Set sbls = CreateSymbols()
sbls.AppendText "012345abcdefg"

Dim pict As stdole.IPictureDisp
Set pict = sbls(0).GetPicture()
```

### 例２．誤り訂正レベルを指定する
CreateSymbols関数の引数に、ErrorCorrectionLevel列挙型の値を設定してSymbolsオブジェクトを生成します。

```VBA
Dim sbls As Symbols
Set sbls = CreateSymbols(ErrorCorrectionLevel.H)
```

### 例３．型番の上限を指定する
CreateSymbols関数の maxVer 引数を設定してSymbolsオブジェクトを生成します。

```VBA
Dim sbls As Symbols
Set sbls = CreateSymbols(maxVer:=10)
```

### 例４．8ビットバイトモードで使用する文字コードを指定する
CreateSymbols関数の charsetName 引数を設定してSymbolsオブジェクトを生成します。

```VBA
Dim sbls As Symbols
Set sbls = CreateSymbols(charsetName:="utf-8")
```

### 例５．分割QRコードを作成する
CreateSymbols関数の引数を設定してSymbolsオブジェクトを生成します。型番の上限を指定しない場合は、型番40を上限として分割されます。  

型番1を超える場合に分割し、各QRコードのIPictureDispオブジェクトを取得する例を示します。

```VBA
Dim sbls As Symbols
Set sbls = CreateSymbols(maxVer:=1, allowStructuredAppend:=True)
sbls.AppendText "abcdefghijklmnopqrstuvwxyz"
    
Dim pict As stdole.IPictureDisp
Dim sbl As Symbol
    
For Each sbl In sbls
    Set pict = sbl.GetPicture()
Next
```

### 例６．BMPファイルへ保存する
SymbolクラスのSaveBitmapメソッドを使用します。

```VBA
Dim sbls As Symbols
Set sbls = CreateSymbols()
sbls.AppendText "012345abcdefg"
    
' 24bpp DIB
sbls(0).SaveBitmap "QRcode.bmp"
    
' 10 pixels per module
sbls(0).SaveBitmap "QRcode.bmp", moduleSize:=10
    
' Specify foreground and background colors.
sbls(0).SaveBitmap "QRcode.bmp", foreRgb:="#0000FF", backRgb:="#FFFF00"
    
' 1bpp DIB
sbls(0).SaveBitmap "QRcode.bmp", monochrome:=True
```

### 例７．SVGファイルへ保存する
SymbolクラスのSaveSvgメソッドを使用します。

```VBA
Dim sbls As Symbols
Set sbls = CreateSymbols()
sbls.AppendText "012345abcdefg"
    
sbls(0).SaveSvg "QRcode.svg"    
```

### 例８．SVGデータを取得する
SymbolクラスのGetSvgメソッドを使用します。

```VBA
Dim sbls As Symbols
Set sbls = CreateSymbols()
sbls.AppendText "012345abcdefg"
    
Dim svg As String
svg = sbls(0).GetSvg()
```

### 例９．クリップボードへ格納する
SymbolクラスのSetToClipboardメソッドを使用します。

```VBA
Dim sbls As Symbols
Set sbls = CreateSymbols()
sbls.AppendText "012345abcdefg"
    
sbls(0).SetToClipBoard
```

