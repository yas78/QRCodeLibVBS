rem ecr ->  L, M(default), Q, H


CScript.exe QRCode.vbs /data:"Hello World" /out:"qr1_truecolor.bmp"
CScript.exe QRCode.vbs /data:"Hello World" /out:"qr1_truecolor.png"
CScript.exe QRCode.vbs /data:"Hello World" /out:"qr1.svg"

CScript.exe QRCode.vbs /data:"Hello World" /out:"qr2_monochrome.bmp" /monochrome:True
CScript.exe QRCode.vbs /data:"Hello World" /out:"qr2_monochrome.png" /monochrome:True

CScript.exe QRCode.vbs /data:"Hello World" /out:"qr3_transparent.png" /transparent:True

CScript.exe QRCode.vbs /data:"Hello World" /out:"qr4_scale10.bmp" /scale:10
CScript.exe QRCode.vbs /data:"Hello World" /out:"qr4_ecrL.bmp" /ecr:L
CScript.exe QRCode.vbs /data:"Hello World" /out:"qr4_forecolor.bmp" /forecolor:#0000ff
CScript.exe QRCode.vbs /data:"Hello World" /out:"qr4_backcolor.bmp" /backcolor:#e0ffff

CScript.exe QRCode.vbs /data:"Hello World" /out:"qr5_scale10.png" /scale:10
CScript.exe QRCode.vbs /data:"Hello World" /out:"qr5_ecrL.png" /ecr:L
CScript.exe QRCode.vbs /data:"Hello World" /out:"qr5_forecolor.png" /forecolor:#0000ff
CScript.exe QRCode.vbs /data:"Hello World" /out:"qr5_backcolor.png" /backcolor:#e0ffff

CScript.exe QRCode.vbs /data:"Hello World" /out:"qr6_all.bmp" /scale:10 /forecolor:#0000ff /backcolor:#e0ffff /ecr:L


CScript.exe QRCode.vbs "test.txt" /out:"qr7_textfile.bmp"

pause
