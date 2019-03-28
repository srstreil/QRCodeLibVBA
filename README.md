# QRCodeLibVBA
QRCodeLibVBAは、Excel VBAで書かれたQRコード生成ライブラリです。  
JIS X 0510に基づくモデル２コードシンボルを生成します。
--------------------------------------------------------------------
(QRCodeLibVBA is a QR code generation library written in Excel VBA.
Generates Model 2 code symbols based on JIS X 0510.)


## 特徴
- 数字・英数字・8ビットバイト・漢字モードに対応しています
- 分割QRコードを作成可能です
- 1bppまたは24bpp BMPファイル(DIB)へ保存可能です
- 1bppまたは24bpp IPictureオブジェクトとして取得可能です  
- 画像の配色(前景色・背景色)を指定可能です
- 8ビットバイトモードでの文字コードを指定可能です
- QRコード画像をクリップボードに保存可能です。
--------------------------------------------------------
(Feature
-Supports numeric, alphanumeric, 8-bit byte and kanji mode
-Can create split QR code
-Save to 1bpp or 24bpp BMP file (DIB)
-Can be obtained as a 1bpp or 24bpp IPicture object
-You can specify the color scheme of the image (foreground color / background color)
-Character code in 8-bit byte mode can be specified
-You can save the QR code image to the clipboard.)

## クイックスタート
32bit版Excelで、QRCodeLib.xlam を参照設定してください。 
--------------------------------------------------------
(quick start
Refer to and set QRCodeLib.xlam in the 32-bit Excel.)

## 使用方法 (Instructions)
### 例１．単一シンボルで構成される(分割QRコードではない)QRコードの、最小限のコードを示します。
--------------------------------------------------------------------------------------
(Example 1. Indicates the minimum code of QR code (not split QR code) composed of a single symbol.)

```vbnet
Public Sub Example()
    Dim sbls As Symbols
    Set sbls = CreateSymbols()
    sbls.AppendText "012345abcdefg"

    Dim pict As stdole.IPicture
    Set pict = sbls(0).Get24bppImage()
    
End Sub
```

### 例２．誤り訂正レベルを指定する
CreateSymbols関数の引数に、ErrorCorrectionLevel列挙型の値を設定してSymbolsオブジェクトを生成します。
--------------------------------------------------------------------------------------------
(Example 2. Specify error correction level
Create a Symbols object by setting the ErrorCorrectionLevel enumeration value to the argument of CreateSymbols function.)

```vbnet
Dim sbls As Symbols
Set sbls = CreateSymbols(ErrorCorrectionLevel.H)
```

### 例３．型番の上限を指定する
CreateSymbols関数の引数を設定してSymbolsオブジェクトを生成します。
-------------------------------------------------------------
(Example 3. Specify upper limit of model number
Set the arguments of CreateSymbols function to create Symbols object.)

```vbnet
Dim sbls As Symbols
Set sbls = CreateSymbols(maxVer:=10)
```

### 例４．8ビットバイトモードで使用する文字コードを指定する
CreateSymbols関数の引数を設定してSymbolsオブジェクトを生成します。
------------------------------------------------------------
(Example 4.8 Specify the character code to be used in bit mode
Set the arguments of CreateSymbols function to create Symbols object.)

```vbnet
Dim sbls As Symbols
Set sbls = CreateSymbols(byteModeCharsetName:="utf-8")
```

### 例５．分割QRコードを作成する
CreateSymbols関数の引数を設定してSymbolsオブジェクトを生成します。型番の上限を指定しない場合は、型番40を上限として分割されます。
-------------------------------------------------------------------------------------------------------------------
(Example 5. Create a split QR code
Set the arguments of CreateSymbols function to create Symbols object. If the upper limit of the model number is not specified, it is divided with the model number 40 as the upper limit.)

```vbnet
Dim sbls As Symbols
Set sbls = CreateSymbols(allowStructuredAppend:=True)
```

型番1を超える場合に分割し、各QRコードのIPictureオブジェクトを取得する例を示します。
-----------------------------------------------------------------------------
(The following shows an example of dividing the model number 1 and acquiring the IPPicture object of each QR code.)

```vbnet
Dim sbls As Symbols
Set sbls = CreateSymbols(maxVer:=1, allowStructuredAppend:=True)
sbls.AppendText "abcdefghijklmnopqrstuvwxyz"

Dim pict As stdole.IPicture
Dim sbl As Symbol

For Each sbl In sbls
    Set pict = sbl.Get24bppImage()
Next
```

### 例６．BMPファイルへ保存する
SymbolクラスのSave1bppDIB、またはSave24bppDIBメソッドを使用します。
--------------------------------------------------------------
(Example 6. Save to BMP file
Use Save1bppDIB or Save24bppDIB method of Symbol class.)

```vbnet
Dim sbls As Symbols
Set sbls = CreateSymbols()
sbls.AppendText "012345abcdefg"

sbls(0).Save1bppDIB "D:\qrcode1bpp1.bmp"
sbls(0).Save1bppDIB "D:\qrcode1bpp2.bmp", 10 ' 10 pixels per module
sbls(0).Save24bppDIB "D:\qrcode24bpp1.bmp"
sbls(0).Save24bppDIB "D:\qrcode24bpp2.bmp", 10 ' 10 pixels per module
```

### 例７．クリップボードへ保存する
SymbolクラスのSetToClipboardメソッドを使用します。
-----------------------------------------------
(Example 7 Save to clipboard
Use the SetToClipboard method of the Symbol class.)

```vbnet
Dim sbls As Symbols
Set sbls = CreateSymbols()
sbls.AppendText "012345abcdefg"

sbls(0).SetToClipboard
sbls(0).SetToClipBoard moduleSize:=10
sbls(0).SetToClipBoard foreRGB:="#0000FF"
sbls(0).SetToClipBoard backRGB:="#00FF00"
```

