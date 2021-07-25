Option Explicit
Include "QRCode.vbs"

Const FORE_COLOR = "#0000ff"
Const BACK_COLOR = "#e0ffff"
Const SCALE = 10

Call Example1
Call Example2


Public Sub Example1()
    Dim sbls: Set sbls = CreateSymbols(ECR_M, 40, False)
    sbls.AppendText "Hello World"
    
    Dim sbl: Set sbl = sbls.Item(0)

    ' BMP truecolor
    sbl.SaveAs "qr_truecolor.bmp"
    ' PNG truecolor
    sbl.SaveAs "qr_truecolor.png"
    ' SVG
    sbl.SaveAs "qr.svg"

    ' BMP monochrome
    sbl.SaveAs2 "qr2_monochrome.bmp", SCALE, True, False, FORE_COLOR, BACK_COLOR
    ' BMP truecolor
    sbl.SaveAs2 "qr2_truecolor.bmp", SCALE, False, False, FORE_COLOR, BACK_COLOR
    ' PNG monochrome
    sbl.SaveAs2 "qr2_monochrome.png", SCALE, True, False, FORE_COLOR, BACK_COLOR
    ' PNG truecolor
    sbl.SaveAs2 "qr2_truecolor.png", SCALE, False, False, FORE_COLOR, BACK_COLOR

    ' PNG transparent
    sbl.SaveAs2 "qr2_transparent.png", SCALE, False, True, "#000000", "#ffffff"

    ' SVG
    sbl.SaveAs2 "qr2_.svg", SCALE, False, False, FORE_COLOR, BACK_COLOR
End Sub


Public Sub Example2()
    Dim sbls: Set sbls = CreateSymbols(ECR_M, 1, True)
    sbls.AppendText "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"

    Dim filename
    Dim i
    For i = 0 To sbls.Count - 1
        filename = "qr_split_" & CStr(i) & ".bmp"
        sbls.Item(i).SaveAs filename
    Next
End Sub


Private Sub Include(ByVal strFile)
    Dim fso: Set fso = WScript.CreateObject("Scripting.FileSystemObject") 
    Dim strDir: strDir = fso.getParentFolderName(WScript.ScriptFullName)
    Dim stream: Set stream = fso.OpenTextFile(strDir & "\" & strFile, 1)

    ExecuteGlobal stream.ReadAll() 
    stream.Close 
End Sub
