# QRCodeLibVBA
QRCodeLibVBAは、Excel VBAで書かれたQRコード生成ライブラリです。  
JIS X 0510に基づくモデル２コードシンボルを生成します。

## 特徴
- 数字・英数字・8ビットバイト・漢字モードに対応しています
- 分割QRコードを作成可能です
- BMP、PNG、SVGファイルに保存可能です
- QRコードをIPictureDispオブジェクトとして取得可能です  
- 配色を指定可能です
- 8ビットバイトモードの文字セットを指定可能です
- QRコードをクリップボードに保存可能です


## クイックスタート
QRCodeLib.xlam を参照設定してください。  


## 使用方法
### 例１．最小限のコードを示します

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

### 例４．8ビットバイトモードの文字セットを指定する
CreateSymbols関数の charsetName 引数を設定してSymbolsオブジェクトを生成します。
（ADODB.Stream に依存しています。使用可能な文字セットはレジストリ[HKEY_CLASSES_ROOT\MIME\Database\Charset]を確認してください。）

```VBA
Dim sbls As Symbols
Set sbls = CreateSymbols(charsetName:="UTF-8")
```

### 例５．分割QRコードを作成する
CreateSymbols関数の引数を設定してSymbolsオブジェクトを生成します。型番の上限を指定しない場合は、型番40を上限に分割されます。  

型番1を上限に分割し、各QRコードのIPictureDispオブジェクトを取得する例を示します。

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

### 例６．ファイルへ保存する
SymbolクラスのSaveAsメソッドを使用します。

```VBA
Dim sbls As Symbols
Set sbls = CreateSymbols()
sbls.AppendText "012345abcdefg"
    
' monochrome BMP
sbls(0).SaveAs "filename"

' true color BMP
sbls(0).SaveAs "filename", fmt:=fmtTrueColor

' monochrome PNG
sbls(0).SaveAs "filename", fmt:=fmtPNG

' true color PNG 
sbls(0).SaveAs "filename", fmt:=fmtPNG + fmtTrueColor

' SVG
sbls(0).SaveAs "filename", fmt:=fmtSVG
        
' 10 pixels per module
sbls(0).SaveAs "filename", moduleSize:=10
    
' specify foreground and background colors
sbls(0).SaveAs "filename", foreRgb:="#0000FF", backRgb:="#FFFF00"
```


### 例７．クリップボードへ保存する
SymbolクラスのSetToClipBoardメソッドを使用します。

```VBA
Dim sbls As Symbols
Set sbls = CreateSymbols()
sbls.AppendText "012345abcdefg"
    
sbls(0).SetToClipBoard
```

