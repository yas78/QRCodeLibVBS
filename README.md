# __QRCodeLibVBS__
QRCodeLibVBSは、VBScriptで書かれたQRコード生成プログラムです。  
JIS X 0510に基づくモデル２コードシンボルを生成します。

## 特徴
- 数字・英数字・8ビットバイト・漢字モードに対応しています
- BMP、SVG、PNGファイルに保存可能です
- 配色を指定可能です
- 分割QRコードを作成可能です

## 使用方法
### 例１．最小限のコードを示します

```bat
rem Command Line
CScript.exe QRCode.vbs /data:"Hello World" /out:"qrcode.bmp"
```

```vbscript
' VBScript
Dim sbls: Set sbls = CreateSymbols(ECR_M, 40, False)
Call sbls.AppendText("Hello World")

Dim sbl: Set sbl = sbls.Item(0)

' BMP truecolor
sbl.SaveAs "qr.bmp"
' PNG truecolor
sbl.SaveAs "qr.png"
' SVG
sbl.SaveAs "qr.svg"
```

### 例２．オプションの使用例
```vbscript
Const FORE_COLOR = "#0000ff"
Const BACK_COLOR = "#e0ffff"
Const SCALE = 10

Dim sbls: Set sbls = CreateSymbols(ECR_M, 40, False)
Call sbls.AppendText("Hello World")

Dim sbl: Set sbl = sbls.Item(0)

' BMP monochrome
sbl.SaveAs2 "qr2_monochrome.bmp", SCALE, True, False, FORE_COLOR, BACK_COLOR
' BMP truecolor
sbl.SaveAs2 "qr2_truecolor.bmp", SCALE, False, False, FORE_COLOR, BACK_COLOR
' PNG monochrome
sbl.SaveAs2 "qr2_monochrome.png", SCALE, True, False, FORE_COLOR, BACK_COLOR
' PNG truecolor
sbl.SaveAs2 "qr2_truecolor.png", SCALE, True, False, FORE_COLOR, BACK_COLOR

' PNG transparent
sbl.SaveAs2 "qr2_transparent.png", SCALE, False, True, FORE_COLOR, BACK_COLOR

' SVG
sbl.SaveAs2 "qr2_.svg", SCALE, False, False, FORE_COLOR, BACK_COLOR
```

### 例３．分割QRコードの作成例
型番1を上限に分割し、各QRコードをファイルに保存する方法を示します。

```vbscript
Dim sbls: Set sbls = CreateSymbols(ECR_M, 1, True)
Call sbls.AppendText("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ")

Dim filename
Dim i
For i = 0 To sbls.Count - 1
    filename = "qr_split" & CStr(i) & ".bmp"
    sbls.Item(i).SaveAs filename
Next
```
