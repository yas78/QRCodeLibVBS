# __QRCodeLibVBS__
QRCodeLibVBSは、VBScriptで書かれたQRコード生成スクリプトです。  
JIS X 0510に基づくモデル２コードシンボルを生成します。

## 特徴
- 数字・英数字・8ビットバイト・漢字モードに対応しています
- 分割QRコードを作成可能です
- 1bppまたは24bpp BMP、SVG形式で保存可能です
- 画像の配色(前景色・背景色)を指定可能です

## 使用方法
### 例１．最小限のコードを示します。
```bat
rem Command Line
CScript.exe QRCode.vbs /data:"Hello World" /out:"qrcode.bmp"
```
```vbscript
' VBScript

Const FORE_COLOR = "#000000"
Const BACK_COLOR = "#FFFFFF"
Const SCALE = 5

Dim sbls: Set sbls = CreateSymbols(ECR_M, 40, False)
Call sbls.AppendText("Hello World")

' 24bpp bitmap
Call sbls.Item(0).Save24bppDIB("qrcode.bmp", SCALE, FORE_COLOR, BACK_COLOR)

' 1bpp bitmap
Call sbls.Item(0).Save1bppDIB("qrcode.bmp", SCALE, FORE_COLOR, BACK_COLOR)

' SVG
Call sbls.Item(0).SaveSvg("qrcode.svg", SCALE, FORE_COLOR)
```

### 例２．分割QRコードの作成例
型番1のデータ量を超える場合に分割し、各QRコードをBMPファイルに保存する方法を示します。
```vbscript
Dim sbls: Set sbls = CreateSymbols(ECR_M, 1, True)
Call sbls.AppendText("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ")

Dim fName
Dim sbl, i
For i = 0 To sbls.Count - 1
    fName = "qrcode" & CStr(i) & ".bmp"
    Call sbls.Item(i).Save24bppDIB(fName, 5, "#000000", "#FFFFFF")
Next
```
