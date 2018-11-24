# QRCodeLibVBA
QRCodeLibVBAは、Excel VBAで書かれたQRコード生成ライブラリです。  
JIS X 0510に基づくモデル２コードシンボルを生成します。

## 特徴
- 数字・英数字・8ビットバイト・漢字モードに対応しています
- 分割QRコードを作成可能です
- 1bppまたは24bpp BMPファイル(DIB)へ保存可能です
- 1bppまたは24bpp IPictureオブジェクトとして取得可能です  
- 画像の配色(前景色・背景色)を指定可能です
- 8ビットバイトモードでの文字コードを指定可能です
- QRコード画像をクリップボードに保存可能です。


## クイックスタート
32bit版Excelで、QRCodeLib.xlam を参照設定してください。  


## 使用方法
### 例１．単一シンボルで構成される(分割QRコードではない)QRコードの、最小限のコードを示します。

```vbnet
Public Sub Example()
    Dim sbls As Symbols
    Set sbls = CreateSymbols()
    sbls.AppendString "012345abcdefg"

    Dim pict As stdole.IPicture
    Set pict = sbls(0).Get24bppImage()
    
End Sub
```

### 例２．誤り訂正レベルを指定する
CreateSymbols関数の引数に、ErrorCorrectionLevel列挙型の値を設定してSymbolsオブジェクトを生成します。

```vbnet
Dim sbls As Symbols
Set sbls = CreateSymbols(ErrorCorrectionLevel.H)
```

### 例３．型番の上限を指定する
CreateSymbols関数の引数を設定してSymbolsオブジェクトを生成します。
```vbnet
Dim sbls As Symbols
Set sbls = CreateSymbols(maxVer:=10)
```

### 例４．8ビットバイトモードで使用する文字コードを指定する
CreateSymbols関数の引数を設定してSymbolsオブジェクトを生成します。
```vbnet
Dim sbls As Symbols
Set sbls = CreateSymbols(byteModeCharsetName:="utf-8")
```

### 例５．分割QRコードを作成する
CreateSymbols関数の引数を設定してSymbolsオブジェクトを生成します。型番の上限を指定しない場合は、型番40を上限として分割されます。
```vbnet
Dim sbls As Symbols
Set sbls = CreateSymbols(allowStructuredAppend:=True)
```

型番1を超える場合に分割し、各QRコードのIPictureオブジェクトを取得する例を示します。

```vbnet
Dim sbls As Symbols
Set sbls = CreateSymbols(maxVer:=1, allowStructuredAppend:=True)
sbls.AppendString "abcdefghijklmnopqrstuvwxyz"

Dim pict As stdole.IPicture
Dim sbl As Symbol

For Each sbl In sbls
    Set pict = sbl.Get24bppImage()
Next
```

### 例６．BMPファイルへ保存する
SymbolクラスのSave1bppDIB、またはSave24bppDIBメソッドを使用します。

```vbnet
Dim sbls As Symbols
Set sbls = CreateSymbols()
sbls.AppendString "012345abcdefg"

sbls(0).Save1bppDIB "D:\qrcode1bpp1.bmp"
sbls(0).Save1bppDIB "D:\qrcode1bpp2.bmp", 10 ' 10 pixels per module
sbls(0).Save24bppDIB "D:\qrcode24bpp1.bmp"
sbls(0).Save24bppDIB "D:\qrcode24bpp2.bmp", 10 ' 10 pixels per module
```

### 例７．クリップボードへ保存する
SymbolクラスのSetToClipboardメソッドを使用します。

```vbnet
Dim sbls As Symbols
Set sbls = CreateSymbols()
sbls.AppendString "012345abcdefg"

sbls(0).SetToClipboard
sbls(0).SetToClipBoard moduleSize:=10
sbls(0).SetToClipBoard foreRGB:="#0000FF"
sbls(0).SetToClipBoard backRGB:="#00FF00"
```

