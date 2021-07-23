rem ecr ->  L, M(default), Q, H
rem colordepth -> 1, 24(default)

CScript.exe QRCode.vbs /data:"Hello World" /out:"qrcode1.bmp"

CScript.exe QRCode.vbs /data:"Hello World" /out:"qrcode2.bmp" /forecolor:#0000FF /backcolor:#E0FFFF /ecr:L /scale:5 /colordepth:1

CScript.exe QRCode.vbs /data:"Hello World" /out:"qrcode3.svg"

CScript.exe QRCode.vbs "test.txt" /out:"qrcode4.bmp"

pause