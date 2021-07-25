Option Explicit

Public Const MIN_VERSION = 1
Public Const MAX_VERSION = 40

Public Const ECR_L = 0
Public Const ECR_M = 1
Public Const ECR_Q = 2
Public Const ECR_H = 3

Private Const MODE_UNKNOWN       = 0
Private Const MODE_NUMERIC       = 1
Private Const MODE_ALPHA_NUMERIC = 2
Private Const MODE_BYTE          = 3
Private Const MODE_KANJI         = 4

Private Const MODEINDICATOR_LENGTH = 4
Private Const MODEINDICATOR_TERMINATOR_VALUE        = &H0
Private Const MODEINDICATOR_NUMERIC_VALUE           = &H1
Private Const MODEINDICATOR_ALPAHNUMERIC_VALUE      = &H2
Private Const MODEINDICATOR_STRUCTURED_APPEND_VALUE = &H3
Private Const MODEINDICATOR_BYTE_VALUE              = &H4
Private Const MODEINDICATOR_KANJI_VALUE             = &H8

Private Const SYMBOLSEQUENCEINDICATOR_POSITION_LENGTH     = 4
Private Const SYMBOLSEQUENCEINDICATOR_TOTAL_NUMBER_LENGTH = 4

Private Const STRUCTUREDAPPEND_PARITY_DATA_LENGTH = 8
Private Const STRUCTUREDAPPEND_HEADER_LENGTH      = 20

Private Const QUIET_ZONE_MIN_WIDTH = 4

Private Const adTypeBinary = 1
Private Const adTypeText   = 2
Private Const adSaveCreateOverWrite = 2

Private Const DIRECTION_UP    = 0
Private Const DIRECTION_DOWN  = 1
Private Const DIRECTION_LEFT  = 2
Private Const DIRECTION_RIGHT = 3

Private Const MIN_MODULE_SIZE = 2

Private Const BLANK         = 0
Private Const WORD          = 1
Private Const ALIGNMENT_PTN = 2
Private Const FINDER_PTN    = 3
Private Const FORMAT_INFO   = 4
Private Const SEPARATOR_PTN = 5
Private Const TIMING_PTN    = 6
Private Const VERSION_INFO  = 7

Private Const CompressionLevel_Fastest = 0
Private Const CompressionLevel_Fast    = 1
Private Const CompressionLevel_Default = 2
Private Const CompressionLevel_Slowest = 3

Private Const DeflateBType_NoCompression = 0
Private Const DeflateBType_CompressedWithFixedHuffmanCodes = 1
Private Const DeflateBType_CompressedWithDynamicHuffmanCodes = 2
Private Const DeflateBType_Reserved = 3

Private Const PngColorType_pGrayscale = 0
Private Const PngColorType_pTrueColor = 2
Private Const PngColorType_pIndexColor = 3
Private Const PngColorType_pGrayscaleAlpha = 4
Private Const PngColorType_pTrueColorAlpha = 6

Private Const PngCompressionMethod_Deflate = 0

Private Const PngFilterType_pNone = 0
Private Const PngFilterType_pSub = 1
Private Const PngFilterType_pUp = 2
Private Const PngFilterType_pAverage = 3
Private Const PngFilterType_pPaeth = 4

Private Const PngInterlaceMethod_pNo = 0
Private Const PngInterlaceMethod_pAdam7 = 1

Private Const BackStyle_bkTransparent = 0
Private Const BackStyle_bkOpaque = 1

Private AlignmentPattern:     Set AlignmentPattern = New AlignmentPattern_
Private CharCountIndicator:   Set CharCountIndicator = New CharCountIndicator_
Private Codeword:             Set Codeword = New Codeword_
Private DataCodeword:         Set DataCodeword = New DataCodeword_
Private FinderPattern:        Set FinderPattern = New FinderPattern_
Private FormatInfo:           Set FormatInfo = New FormatInfo_
Private GaloisField256:       Set GaloisField256 = New GaloisField256_
Private GeneratorPolynomials: Set GeneratorPolynomials = New GeneratorPolynomials_
Private Masking:              Set Masking = New Masking_
Private MaskingPenaltyScore:  Set MaskingPenaltyScore = New MaskingPenaltyScore_
Private Module:               Set Module = New Module_
Private QuietZone:            Set QuietZone = New QuietZone_
Private RemainderBit:         Set RemainderBit = New RemainderBit_
Private RSBlock:              Set RSBlock = New RSBlock_
Private Separator:            Set Separator = New Separator_
Private TimingPattern:        Set TimingPattern = New TimingPattern_
Private VersionInfo:          Set VersionInfo = New VersionInfo_
Private ColorCode:            Set ColorCode = New ColorCode_
Private GraphicPath:          Set GraphicPath = New GraphicPath_
Private DIB:                  Set DIB = New DIB_
Private PNG:                  Set PNG = New PNG_
Private CRC32:                Set CRC32 = New CRC32_
Private ADLER32:              Set ADLER32 = New ADLER32_
Private Deflate:              Set Deflate = New Deflate_
Private ZLIB:                 Set ZLIB = New ZLIB_
Private BitConverter:         Set BitConverter = New BitConverter_


Call Main(WScript.Arguments)


Public Sub Main(ByVal args)
    If args.Count = 0 Then Exit Sub

    Dim params
    Set params = GetParams(args)
    If params Is Nothing Then
        Call WScript.Quit(-1)
    End If

    If Len(params("data")) = 0 Then
        Call WScript.Quit(-1)
    End If

    Dim sbls: Set sbls = CreateSymbols(CLng(params("ecr")), MAX_VERSION, False)
    Call sbls.AppendText(params("data"))

    Dim sbl: Set sbl = sbls.Item(0)
    Call sbl.SaveAs2( _
        params("out"), CLng(params("scale")), CBool(params("monochrome")), _
        CBool(params("transparent")), params("forecolor"), params("backcolor"))

    Call WScript.Quit(0)
End Sub

Private Function GetParams(ByVal args)
    Dim ks
    ks = Array("data", "out", "monochrome", "transparent", "forecolor", "backcolor", "ecr", "scale")

    Dim params
    Set params = CreateObject("Scripting.Dictionary")
    Dim k, v

    For Each k In ks
        Call params.Add(k, Empty)
    Next

    Dim fso, ts
    Set fso = CreateObject("Scripting.FileSystemObject")

    If args.UnNamed.Count > 0 Then
        If Not fso.FileExists(args.UnNamed(0)) Then
            Call WScript.Echo("file not found")
            Exit Function
        End If

        Set ts = fso.OpenTextFile(args.UnNamed(0))
        params("data") = ts.ReadAll()
        ts.Close
    End If

    params("scale") = 5
    params("forecolor") = ColorCode.BLACK
    params("backcolor") = ColorCode.WHITE
    params("monochrome") = False
    params("transparent") = False
    params("ecr") = "M"

    For Each k In ks
        If args.Named.Exists(k) Then
            v = args.Named.Item(k)
            If Len(v) = 0 Then
                Call WScript.Echo("argument error [" & k  & "]")
                Exit Function
            End If
            If IsNumeric(v) Then
                v = CLng(v)
            End If
            params(k) = v
        End IF
    Next

    If Len(params("out")) = 0 Then
        Call WScript.Echo("argument error [out]")
        Exit Function
    End If

    Select Case LCase(CStr(params("monochrome")))
        Case "true", "false"
            ' NOP
        Case Else
            Call WScript.Echo("argument error [monochrome]")
            Exit Function
    End Select

    Select Case LCase(CStr(params("transparent")))
        Case "true", "false"
            ' NOP
        Case Else
            Call WScript.Echo("argument error [transparent]")
            Exit Function
    End Select

    If Not ColorCode.IsWebColor(params("forecolor")) Then
        Call WScript.Echo("argument error [forecolor]")
        Exit Function
    End If

    If Not ColorCode.IsWebColor(params("backcolor")) Then
        Call WScript.Echo("argument error 'backcolor'")
        Exit Function
    End If

    If Not IsNumeric(params("scale")) Then
        Call WScript.Echo("argument error [scale]")
        Exit Function
    End If

    If params("scale") < MIN_MODULE_SIZE Then
        Call WScript.Echo("argument error [scale]")
        Exit Function
    End If

    Select Case UCase(params("ecr"))
        Case "L"
            v = ECR_L
        Case "M"
            v = ECR_M
        Case "Q"
            v = ECR_Q
        Case "H"
            v = ECR_H
        Case Else
            Call WScript.Echo("argument error [ecr]")
            Exit Function
    End Select
    params("ecr") = v

    Dim ext: ext = fso.GetExtensionName(params("out"))
    Select Case LCase(ext)
        Case "bmp", "svg", "png"
            ' NOP
        Case Else
            Call WScript.Echo("argument error [out] unsupported file type")
            Exit Function
    End Select

    Set GetParams = params
End Function

Public Function CreateSymbols(ByVal ecLevel, ByVal maxVer, ByVal allowStructuredAppend)
    Select Case ecLevel
        Case ECR_L ,ECR_M, ECR_Q, ECR_H
            ' NOP
        Case Else
            Call Err.Raise(5)
    End Select

    If Not (MIN_VERSION <= maxVer And maxVer <= MAX_VERSION) Then Call Err.Raise(5)

    Dim ret
    Set ret = New Symbols
    Call ret.Init(ecLevel, maxVer, allowStructuredAppend)

    Set CreateSymbols = ret
End Function

Private Function CreateEncoder(ByVal encMode)
    Dim ret

    Select Case encMode
        Case MODE_NUMERIC
            Set ret = New NumericEncoder
        Case MODE_ALPHA_NUMERIC
            Set ret = New AlphanumericEncoder
        Case MODE_BYTE
            Set ret = New ByteEncoder
        Case MODE_KANJI
            Set ret = New KanjiEncoder
        Case Else
            Call Err.Raise(5)
    End Select

    Set CreateEncoder = ret
End Function

Private Function IsDark(ByVal arg)
    IsDark = arg > BLANK
End Function


Class AlignmentPattern_
    Private m_lst(40)

    Private Sub Class_Initialize()
        m_lst(2)  = Array(6, 18)
        m_lst(3)  = Array(6, 22)
        m_lst(4)  = Array(6, 26)
        m_lst(5)  = Array(6, 30)
        m_lst(6)  = Array(6, 34)
        m_lst(7)  = Array(6, 22, 38)
        m_lst(8)  = Array(6, 24, 42)
        m_lst(9)  = Array(6, 26, 46)
        m_lst(10) = Array(6, 28, 50)
        m_lst(11) = Array(6, 30, 54)
        m_lst(12) = Array(6, 32, 58)
        m_lst(13) = Array(6, 34, 62)
        m_lst(14) = Array(6, 26, 46, 66)
        m_lst(15) = Array(6, 26, 48, 70)
        m_lst(16) = Array(6, 26, 50, 74)
        m_lst(17) = Array(6, 30, 54, 78)
        m_lst(18) = Array(6, 30, 56, 82)
        m_lst(19) = Array(6, 30, 58, 86)
        m_lst(20) = Array(6, 34, 62, 90)
        m_lst(21) = Array(6, 28, 50, 72, 94)
        m_lst(22) = Array(6, 26, 50, 74, 98)
        m_lst(23) = Array(6, 30, 54, 78, 102)
        m_lst(24) = Array(6, 28, 54, 80, 106)
        m_lst(25) = Array(6, 32, 58, 84, 110)
        m_lst(26) = Array(6, 30, 58, 86, 114)
        m_lst(27) = Array(6, 34, 62, 90, 118)
        m_lst(28) = Array(6, 26, 50, 74, 98, 122)
        m_lst(29) = Array(6, 30, 54, 78, 102, 126)
        m_lst(30) = Array(6, 26, 52, 78, 104, 130)
        m_lst(31) = Array(6, 30, 56, 82, 108, 134)
        m_lst(32) = Array(6, 34, 60, 86, 112, 138)
        m_lst(33) = Array(6, 30, 58, 86, 114, 142)
        m_lst(34) = Array(6, 34, 62, 90, 118, 146)
        m_lst(35) = Array(6, 30, 54, 78, 102, 126, 150)
        m_lst(36) = Array(6, 24, 50, 76, 102, 128, 154)
        m_lst(37) = Array(6, 28, 54, 80, 106, 132, 158)
        m_lst(38) = Array(6, 32, 58, 84, 110, 136, 162)
        m_lst(39) = Array(6, 26, 54, 82, 110, 138, 166)
        m_lst(40) = Array(6, 30, 58, 86, 114, 142, 170)
    End Sub

    Public Sub Place(ByVal ver, ByRef moduleMatrix())
        Dim VAL
        VAL = ALIGNMENT_PTN

        Dim centerArray
        centerArray = m_lst(ver)

        Dim maxIndex
        maxIndex = UBound(centerArray)

        Dim i, j
        Dim r, c

        For i = 0 To maxIndex
            r = centerArray(i)

            For j = 0 To maxIndex
                c = centerArray(j)

                If (i = 0 And j = 0 Or _
                    i = 0 And j = maxIndex Or _
                    i = maxIndex And j = 0) = False Then

                    moduleMatrix(r - 2)(c - 2) = VAL
                    moduleMatrix(r - 2)(c - 1) = VAL
                    moduleMatrix(r - 2)(c + 0) = VAL
                    moduleMatrix(r - 2)(c + 1) = VAL
                    moduleMatrix(r - 2)(c + 2) = VAL

                    moduleMatrix(r - 1)(c - 2) = VAL
                    moduleMatrix(r - 1)(c - 1) = -VAL
                    moduleMatrix(r - 1)(c + 0) = -VAL
                    moduleMatrix(r - 1)(c + 1) = -VAL
                    moduleMatrix(r - 1)(c + 2) = VAL

                    moduleMatrix(r + 0)(c - 2) = VAL
                    moduleMatrix(r + 0)(c - 1) = -VAL
                    moduleMatrix(r + 0)(c + 0) = VAL
                    moduleMatrix(r + 0)(c + 1) = -VAL
                    moduleMatrix(r + 0)(c + 2) = VAL

                    moduleMatrix(r + 1)(c - 2) = VAL
                    moduleMatrix(r + 1)(c - 1) = -VAL
                    moduleMatrix(r + 1)(c + 0) = -VAL
                    moduleMatrix(r + 1)(c + 1) = -VAL
                    moduleMatrix(r + 1)(c + 2) = VAL

                    moduleMatrix(r + 2)(c - 2) = VAL
                    moduleMatrix(r + 2)(c - 1) = VAL
                    moduleMatrix(r + 2)(c + 0) = VAL
                    moduleMatrix(r + 2)(c + 1) = VAL
                    moduleMatrix(r + 2)(c + 2) = VAL
                End If
            Next
        Next
    End Sub
End Class


Class AlphanumericEncoder
    Private m_data
    Private m_charCounter
    Private m_bitCounter

    Private m_encNumeric

    Private Sub Class_Initialize()
        m_data = Empty
        m_charCounter = 0
        m_bitCounter = 0

        Set m_encNumeric = New NumericEncoder
    End Sub

    Public Property Get BitCount()
        BitCount = m_bitCounter
    End Property

    Public Property Get CharCount()
        CharCount = m_charCounter
    End Property

    Public Property Get EncodingMode()
        EncodingMode = MODE_ALPHA_NUMERIC
    End Property

    Public Property Get ModeIndicator()
        ModeIndicator = MODEINDICATOR_ALPAHNUMERIC_VALUE
    End Property

    Public Function Append(ByVal c)
        Dim wd
        wd = ConvertCharCode(c)

        If m_charCounter Mod 2 = 0 Then
            If m_charCounter = 0 Then
                ReDim m_data(0)
            Else
                ReDim Preserve m_data(UBound(m_data) + 1)
            End If

            m_data(UBound(m_data)) = wd
        Else
            m_data(UBound(m_data)) = m_data(UBound(m_data)) * 45
            m_data(UBound(m_data)) = m_data(UBound(m_data)) + wd
        End If

        Dim ret
        ret = GetCodewordBitLength(c)
        m_bitCounter = m_bitCounter + ret
        m_charCounter = m_charCounter + 1

        Append = ret
    End Function

    Public Function GetCodewordBitLength(ByVal c)
        If m_charCounter Mod 2 = 0 Then
            GetCodewordBitLength = 6
        Else
            GetCodewordBitLength = 5
        End If
    End Function

    Public Function GetBytes()
        Dim bs
        Set bs = New BitSequence

        Dim bitLength
        bitLength = 11

        Dim i
        For i = 0 To UBound(m_data) - 1
            Call bs.Append(m_data(i), bitLength)
        Next

        If m_charCounter Mod 2 = 0 Then
            bitLength = 11
        Else
            bitLength = 6
        End If

        Call bs.Append(m_data(UBound(m_data)), bitLength)

        GetBytes = bs.GetBytes()
    End Function

    Public Function ConvertCharCode(ByVal c)
        Dim code
        code = Asc(c)

        ' (Space)
        If code = 32 Then
            ConvertCharCode = 36
        ' $ %
        ElseIf code = 36 Or code = 37 Then
            ConvertCharCode = code + 1
        ' * +
        ElseIf code = 42 Or code = 43 Then
            ConvertCharCode = code - 3
        ' - .
        ElseIf code = 45 Or code = 46 Then
            ConvertCharCode = code - 4
        ' /
        ElseIf code = 47 Then
            ConvertCharCode = 43
        ' 0 - 9
        ElseIf 48 <= code And code <= 57 Then
            ConvertCharCode = code - 48
        ' :
        ElseIf code = 58 Then
            ConvertCharCode = 44
        ' A - Z
        ElseIf 65 <= code And code <= 90 Then
            ConvertCharCode = code - 55
        Else
            ConvertCharCode = -1
        End If
    End Function

    Public Function InSubset(ByVal c)
        InSubset = ConvertCharCode(c) > -1
    End Function

    Public Function InExclusiveSubset(ByVal c)
        If m_encNumeric.InSubset(c) Then
            InExclusiveSubset = False
            Exit Function
        End If

        InExclusiveSubset = InSubset(c)
    End Function
End Class


Class BinaryWriter
    Private m_byteTable(255)
    Private m_stream

    Private Sub Class_Initialize()
        Call MakeByteTable

        Set m_stream = CreateObject("ADODB.Stream")
        m_stream.Type = adTypeBinary
        Call m_stream.Open
    End Sub

    Public Property Get Stream()
        Set Stream = m_stream
    End Property

    Public Property Get Size()
        Size = m_stream.Size
    End Property

    Private Sub MakeByteTable()
        Dim sr
        Set sr = CreateObject("ADODB.Stream")
        sr.Type = adTypeText
        sr.Charset = "unicode"
        Call sr.Open

        Dim i
        For i = 0 To 255
            sr.WriteText ChrW(i)
        Next

        sr.Position = 0
        sr.Type = adTypeBinary
        sr.Position = 2

        For i = 0 To 255
            m_byteTable(i) = sr.Read(1)
            sr.Position = sr.Position + 1
        Next

        Call sr.Close
    End Sub

    Public Sub Append(ByVal arg)
        If (VarType(arg) And vbArray) = 0 Then
            arg = Array(arg)
        End If

        Dim temp
        Dim v

        For Each v In arg
            Select Case VarType(v)
                Case vbByte
                    Call m_stream.Write(m_byteTable(v))
                Case vbInteger
                    temp = v And &HFF&
                    Call m_stream.Write(m_byteTable(temp))

                    temp = (v And &HFF00&) \ 2 ^ 8
                    Call m_stream.Write(m_byteTable(temp))
                Case vbLong
                    temp = v And &HFF&
                    Call m_stream.Write(m_byteTable(temp))
                    temp = (v And &HFF00&) \ 2 ^ 8
                    Call m_stream.Write(m_byteTable(temp))
                    temp = (v And &HFF0000) \ 2 ^ 16
                    Call m_stream.Write(m_byteTable(temp))

                    If (v And &H80000000) <> 0 Then
                        temp = ((v And &H7F000000) \ 2 ^ 24) Or &H80
                    Else
                        temp = v \ 2 ^ 24
                    End If
                    Call m_stream.Write(m_byteTable(temp))
                Case Else
                    Call Err.Raise(5)
            End Select
        Next
    End Sub

    Public Sub CopyTo(ByVal destBinaryWriter)
        m_stream.Position = 0
        Call m_stream.CopyTo(destBinaryWriter.Stream)
    End Sub

    Public Sub SaveToFile(ByVal FileName, ByVal SaveOptions)
        Call m_stream.SaveToFile(FileName, SaveOptions)
    End Sub
End Class


Class BitSequence
    Private m_buffer()
    Private m_bitCounter
    Private m_space
    Private m_index

    Private Sub Class_Initialize()
        Call Clear
    End Sub

    Public Property Get Length()
        Length = m_bitCounter
    End Property

    Public Sub Clear()
        Erase m_buffer
        m_index = -1
        m_bitCounter = 0
        m_space = 0
    End Sub

    Public Sub Append(ByVal data, ByVal bitLength)
        Dim remainingLength
        remainingLength = bitLength
        Dim remainingData
        remainingData = data

        Dim temp

        Do While remainingLength > 0
            If m_space = 0 Then
                m_space = 8
                m_index = m_index + 1
                ReDim Preserve m_buffer(m_index)
            End If

            temp = m_buffer(m_index)

            If m_space < remainingLength Then
                temp = CByte(temp Or remainingData \ (2 ^ (remainingLength - m_space)))
                remainingData = remainingData And ((2 ^ (remainingLength - m_space)) - 1)

                m_bitCounter = m_bitCounter + m_space
                remainingLength = remainingLength - m_space
                m_space = 0
            Else
                temp = CByte(temp Or remainingData * (2 ^ (m_space - remainingLength)))
                m_bitCounter = m_bitCounter + remainingLength
                m_space = m_space - remainingLength
                remainingLength = 0
            End If

            m_buffer(m_index) = temp
        Loop
    End Sub

    Public Function GetBytes()
        If m_index < 0 Then Call Err.Raise(51)
        GetBytes = m_buffer
    End Function
End Class


Class ByteEncoder
    Private m_data
    Private m_charCounter
    Private m_bitCounter

    Private m_encAlpha
    Private m_encKanji

    Private Sub Class_Initialize()
        m_data = Empty
        m_charCounter = 0
        m_bitCounter = 0

        Set m_encAlpha = New AlphanumericEncoder
        Set m_encKanji = New KanjiEncoder
    End Sub

    Public Property Get BitCount()
        BitCount = m_bitCounter
    End Property

    Public Property Get CharCount()
        CharCount = m_charCounter
    End Property

    Public Property Get EncodingMode()
        EncodingMode = MODE_BYTE
    End Property

    Public Property Get ModeIndicator()
        ModeIndicator = MODEINDICATOR_BYTE_VALUE
    End Property

    Public Function Append(ByVal c)
        If m_charCounter = 0 Then
            ReDim m_data(0)
        Else
            ReDim Preserve m_data(UBound(m_data) + 1)
        End If

        Dim wd
        wd = Asc(c) And &HFFFF&
        m_data(UBound(m_data)) = wd

        Dim ret
        ret = GetCodewordBitLength(c)
        m_bitCounter = m_bitCounter + ret
        m_charCounter = m_charCounter + (ret \ 8)

        Append = ret
    End Function

    Public Function GetCodewordBitLength(ByVal c)
        Dim code
        code = Asc(c) And &HFFFF&

        If code > &HFF Then
            GetCodewordBitLength = 16
        Else
            GetCodewordBitLength = 8
        End If
    End Function

    Public Function GetBytes()
        GetBytes = m_data
    End Function

    Public Function InSubset(ByVal c)
        InSubset = True
    End Function

    Public Function InExclusiveSubset(ByVal c)
        If m_encAlpha.InSubset(c) Then
            InExclusiveSubset = False
            Exit Function
        End If

        If m_encKanji.InSubset(c) Then
            InExclusiveSubset = False
            Exit Function
        End If

        InExclusiveSubset = InSubset(c)
    End Function
End Class


Class CharCountIndicator_
    Public Function GetLength(ByVal ver, ByVal encMode)
        If 1 <= ver And ver <= 9 Then
            Select Case encMode
                Case MODE_NUMERIC
                    GetLength = 10
                Case MODE_ALPHA_NUMERIC
                    GetLength = 9
                Case MODE_BYTE
                    GetLength = 8
                Case MODE_KANJI
                    GetLength = 8
                Case Else
                    Call Err.Raise(5)
            End Select
        ElseIf 10 <= ver And ver <= 26 Then
            Select Case encMode
                Case MODE_NUMERIC
                    GetLength = 12
                Case MODE_ALPHA_NUMERIC
                    GetLength = 11
                Case MODE_BYTE
                    GetLength = 16
                Case MODE_KANJI
                    GetLength = 10
                Case Else
                    Call Err.Raise(5)
            End Select
        ElseIf 27 <= ver And ver <= 40 Then
            Select Case encMode
                Case MODE_NUMERIC
                    GetLength = 14
                Case MODE_ALPHA_NUMERIC
                    GetLength = 13
                Case MODE_BYTE
                    GetLength = 16
                Case MODE_KANJI
                    GetLength = 12
                Case Else
                    Call Err.Raise(5)
            End Select
        Else
            Call Err.Raise(5)
        End If
    End Function
End Class


Class Codeword_
    Private m_totalNumbers

    Private Sub Class_Initialize()
        m_totalNumbers = Array( _
              -1, _
              26,   44,   70,  100,  134,  172,  196,  242,  292,  346, _
             404,  466,  532,  581,  655,  733,  815,  901,  991, 1085, _
            1156, 1258, 1364, 1474, 1588, 1706, 1828, 1921, 2051, 2185, _
            2323, 2465, 2611, 2761, 2876, 3034, 3196, 3362, 3532, 3706 _
        )
    End Sub

    Public Function GetTotalNumber(ByVal ver)
        GetTotalNumber = m_totalNumbers(ver)
    End Function
End Class


Class DataCodeword_
    Private m_totalNumbers

    Private Sub Class_Initialize()
        Dim ecLevelL
        ecLevelL = Array( _
               0, _
              19,   34,   55,   80,  108,  136,  156,  194,  232,  274, _
             324,  370,  428,  461,  523,  589,  647,  721,  795,  861, _
             932, 1006, 1094, 1174, 1276, 1370, 1468, 1531, 1631, 1735, _
            1843, 1955, 2071, 2191, 2306, 2434, 2566, 2702, 2812, 2956 _
        )

        Dim ecLevelM
        ecLevelM = Array( _
               0, _
              16,   28,   44,   64,   86,  108,  124,  154,  182,  216, _
             254,  290,  334,  365,  415,  453,  507,  563,  627,  669, _
             714,  782,  860,  914, 1000, 1062, 1128, 1193, 1267, 1373, _
            1455, 1541, 1631, 1725, 1812, 1914, 1992, 2102, 2216, 2334 _
        )

        Dim ecLevelQ
        ecLevelQ = Array( _
               0, _
              13,   22,   34,   48,   62,   76,   88,  110,  132,  154, _
             180,  206,  244,  261,  295,  325,  367,  397,  445,  485, _
             512,  568,  614,  664,  718,  754,  808,  871,  911,  985, _
            1033, 1115, 1171, 1231, 1286, 1354, 1426, 1502, 1582, 1666 _
        )

        Dim ecLevelH
        ecLevelH = Array( _
              0, _
              9,  16,  26,  36,  46,   60,   66,   86,  100,  122, _
            140, 158, 180, 197, 223,  253,  283,  313,  341,  385, _
            406, 442, 464, 514, 538,  596,  628,  661,  701,  745, _
            793, 845, 901, 961, 986, 1054, 1096, 1142, 1222, 1276 _
        )

        m_totalNumbers = Array(ecLevelL, ecLevelM, ecLevelQ, ecLevelH)
    End Sub

    Public Function GetTotalNumber(ByVal ecLevel, ByVal ver)
        GetTotalNumber = m_totalNumbers(ecLevel)(ver)
    End Function
End Class


Class FinderPattern_
    Private m_finderPattern

    Private Sub Class_Initialize()
        Dim VAL
        VAL = FINDER_PTN

        m_finderPattern = Array( _
            Array(VAL,  VAL,  VAL,  VAL,  VAL,  VAL,  VAL), _
            Array(VAL, -VAL, -VAL, -VAL, -VAL, -VAL,  VAL), _
            Array(VAL, -VAL,  VAL,  VAL,  VAL, -VAL,  VAL), _
            Array(VAL, -VAL,  VAL,  VAL,  VAL, -VAL,  VAL), _
            Array(VAL, -VAL,  VAL,  VAL,  VAL, -VAL,  VAL), _
            Array(VAL, -VAL, -VAL, -VAL, -VAL, -VAL,  VAL), _
            Array(VAL,  VAL,  VAL,  VAL,  VAL,  VAL,  VAL) _
        )
    End Sub

    Public Sub Place(ByRef moduleMatrix())
        Dim offset
        offset = (UBound(moduleMatrix) + 1) - (UBound(m_finderPattern) + 1)

        Dim i, j
        Dim v

        For i = 0 To UBound(m_finderPattern)
            For j = 0 To UBound(m_finderPattern(i))
                v = m_finderPattern(i)(j)

                moduleMatrix(i)(j) = v
                moduleMatrix(i)(j + offset) = v
                moduleMatrix(i + offset)(j) = v
            Next
        Next
    End Sub
End Class


Class FormatInfo_
    Private VAL
    Private m_formatInfoValues
    Private m_formatInfoMaskArray

    Private Sub Class_Initialize()
        VAL = FORMAT_INFO
        m_formatInfoValues = Array( _
               &H0&,  &H537&,  &HA6E&,  &HF59&, &H11EB&, &H14DC&, &H1B85&, &H1EB2&, &H23D6&, &H26E1&, _
            &H29B8&, &H2C8F&, &H323D&, &H370A&, &H3853&, &H3D64&, &H429B&, &H47AC&, &H48F5&, &H4DC2&, _
            &H5370&, &H5647&, &H591E&, &H5C29&, &H614D&, &H647A&, &H6B23&, &H6E14&, &H70A6&, &H7591&, _
            &H7AC8&, &H7FFF& _
        )

        m_formatInfoMaskArray = Array(0, 1, 0, 0, 1, 0, 0, 0, 0, 0, 1, 0, 1, 0, 1)
    End Sub

    Public Sub Place(ByVal ecLevel, ByVal maskPatternReference, ByRef moduleMatrix())
        Dim formatInfoValue
        formatInfoValue = GetFormatInfoValue(ecLevel, maskPatternReference)

        Dim temp
        Dim v

        Dim i

        Dim r1
        r1 = 0

        Dim c1
        c1 = UBound(moduleMatrix)

        For i = 0 To 7
            If (formatInfoValue And (2 ^ i)) > 0 Then
                temp = 1 Xor m_formatInfoMaskArray(i)
            Else
                temp = 0 Xor m_formatInfoMaskArray(i)
            End If

            If temp > 0 Then
                v = VAL
            Else
                v = -VAL
            End IF

            moduleMatrix(r1)(8) = v
            moduleMatrix(8)(c1) = v

            r1 = r1 + 1
            c1 = c1 - 1

            If r1 = 6 Then
                r1 = r1 + 1
            End If
        Next

        Dim r2
        r2 = UBound(moduleMatrix) - 6

        Dim c2
        c2 = 7

        For i = 8 To 14
            If (formatInfoValue And (2 ^ i)) > 0 Then
                temp = 1 Xor m_formatInfoMaskArray(i)
            Else
                temp = 0 Xor m_formatInfoMaskArray(i)
            End If

            If temp > 0 Then
                v = VAL
            Else
                v = -VAL
            End IF

            moduleMatrix(r2)(8) = v
            moduleMatrix(8)(c2) = v

            r2 = r2 + 1
            c2 = c2 - 1

            If c2 = 6 Then
                c2 = c2 - 1
            End If
        Next

        moduleMatrix(UBound(moduleMatrix) - 7)(8) = VAL
    End Sub

    Public Sub PlaceTempBlank(ByRef moduleMatrix())
        Dim i
        For i = 0 To 8
            If i <> 6 Then
                moduleMatrix(8)(i) = -VAL
                moduleMatrix(i)(8) = -VAL
            End If
        Next

        Dim numModulesPerSide
        numModulesPerSide = UBound(moduleMatrix) + 1

        For i = UBound(moduleMatrix) - 7 To UBound(moduleMatrix)
            moduleMatrix(8)(i) = -VAL
            moduleMatrix(i)(8) = -VAL
        Next

        moduleMatrix(UBound(moduleMatrix) - 7)(8) = -VAL
    End Sub

    Private Function GetFormatInfoValue(ByVal ecLevel, ByVal maskPatternReference)
        Dim indicator

        Select Case ecLevel
            Case ECR_L
                indicator = 1
            Case ECR_M
                indicator = 0
            Case ECR_Q
                indicator = 3
            Case ECR_H
                indicator = 2
            Case Else
                Call Err.Raise(5)
        End Select

        GetFormatInfoValue = m_formatInfoValues((indicator * 2 ^ 3) Or maskPatternReference)
    End Function
End Class


Class GaloisField256_
    Private m_intToExpTable
    Private m_expToIntTable

    Private Sub Class_Initialize()
        m_intToExpTable = Array( _
             -1,   0,   1,  25,   2,  50,  26, 198,   3, 223,  51, 238,  27, 104, 199,  75, _
              4, 100, 224,  14,  52, 141, 239, 129,  28, 193, 105, 248, 200,   8,  76, 113, _
              5, 138, 101,  47, 225,  36,  15,  33,  53, 147, 142, 218, 240,  18, 130,  69, _
             29, 181, 194, 125, 106,  39, 249, 185, 201, 154,   9, 120,  77, 228, 114, 166, _
              6, 191, 139,  98, 102, 221,  48, 253, 226, 152,  37, 179,  16, 145,  34, 136, _
             54, 208, 148, 206, 143, 150, 219, 189, 241, 210,  19,  92, 131,  56,  70,  64, _
             30,  66, 182, 163, 195,  72, 126, 110, 107,  58,  40,  84, 250, 133, 186,  61, _
            202,  94, 155, 159,  10,  21, 121,  43,  78, 212, 229, 172, 115, 243, 167,  87, _
              7, 112, 192, 247, 140, 128,  99,  13, 103,  74, 222, 237,  49, 197, 254,  24, _
            227, 165, 153, 119,  38, 184, 180, 124,  17,  68, 146, 217,  35,  32, 137,  46, _
             55,  63, 209,  91, 149, 188, 207, 205, 144, 135, 151, 178, 220, 252, 190,  97, _
            242,  86, 211, 171,  20,  42,  93, 158, 132,  60,  57,  83,  71, 109,  65, 162, _
             31,  45,  67, 216, 183, 123, 164, 118, 196,  23,  73, 236, 127,  12, 111, 246, _
            108, 161,  59,  82,  41, 157,  85, 170, 251,  96, 134, 177, 187, 204,  62,  90, _
            203,  89,  95, 176, 156, 169, 160,  81,  11, 245,  22, 235, 122, 117,  44, 215,  _
             79, 174, 213, 233, 230, 231, 173, 232, 116, 214, 244, 234, 168,  80,  88, 175 _
        )

        m_expToIntTable = Array( _
              1,   2,   4,   8,  16,  32,  64, 128,  29,  58, 116, 232, 205, 135,  19,  38, _
             76, 152,  45,  90, 180, 117, 234, 201, 143,   3,   6,  12,  24,  48,  96, 192, _
            157,  39,  78, 156,  37,  74, 148,  53, 106, 212, 181, 119, 238, 193, 159,  35, _
             70, 140,   5,  10,  20,  40,  80, 160,  93, 186, 105, 210, 185, 111, 222, 161, _
             95, 190,  97, 194, 153,  47,  94, 188, 101, 202, 137,  15,  30,  60, 120, 240, _
            253, 231, 211, 187, 107, 214, 177, 127, 254, 225, 223, 163,  91, 182, 113, 226, _
            217, 175,  67, 134,  17,  34,  68, 136,  13,  26,  52, 104, 208, 189, 103, 206, _
            129,  31,  62, 124, 248, 237, 199, 147,  59, 118, 236, 197, 151,  51, 102, 204, _
            133,  23,  46,  92, 184, 109, 218, 169,  79, 158,  33,  66, 132,  21,  42,  84, _
            168,  77, 154,  41,  82, 164,  85, 170,  73, 146,  57, 114, 228, 213, 183, 115, _
            230, 209, 191,  99, 198, 145,  63, 126, 252, 229, 215, 179, 123, 246, 241, 255, _
            227, 219, 171,  75, 150,  49,  98, 196, 149,  55, 110, 220, 165,  87, 174,  65, _
            130,  25,  50, 100, 200, 141,   7,  14,  28,  56, 112, 224, 221, 167,  83, 166, _
             81, 162,  89, 178, 121, 242, 249, 239, 195, 155,  43,  86, 172,  69, 138,   9, _
             18,  36,  72, 144,  61, 122, 244, 245, 247, 243, 251, 235, 203, 139,  11,  22, _
             44,  88, 176, 125, 250, 233, 207, 131,  27,  54, 108, 216, 173,  71, 142,   1 _
        )
    End Sub

    Public Function ToExp(ByVal arg)
        ToExp = m_intToExpTable(arg)
    End Function

    Public Function ToInt(ByVal arg)
        ToInt = m_expToIntTable(arg)
    End Function
End Class


Class GeneratorPolynomials_
    Private m_gp(68)

    Private Sub Class_Initialize()
        m_gp(7)  = Array( 21, 102, 238, 149, 146, 229,  87,   0)
        m_gp(10) = Array( 45,  32,  94,  64,  70, 118,  61,  46,  67, 251,   0)
        m_gp(13) = Array( 78, 140, 206, 218, 130, 104, 106, 100,  86, 100, 176, 152,  74,   0)
        m_gp(15) = Array(105,  99,   5, 124, 140, 237,  58,  58,  51,  37, 202,  91,  61, 183,   8,   0)
        m_gp(16) = Array(120, 225, 194, 182, 169, 147, 191,  91,   3,  76, 161, 102, 109, 107, 104, 120,   0)
        m_gp(17) = Array(136, 163, 243,  39, 150,  99,  24, 147, 214, 206, 123, 239,  43,  78, 206, 139,  43,   0)
        m_gp(18) = Array(153,  96,  98,   5, 179, 252, 148, 152, 187,  79, 170, 118,  97, 184,  94, 158, 234, 215,   0)
        m_gp(20) = Array(190, 188, 212, 212, 164, 156, 239,  83, 225, 221, 180, 202, 187,  26, 163,  61,  50,  79,  60,  17,   0)
        m_gp(22) = Array(231, 165, 105, 160, 134, 219,  80,  98, 172,   8,  74, 200,  53, 221, 109,  14, 230,  93, 242, 247, 171, 210,   0)
        m_gp(24) = Array( 21, 227,  96,  87, 232, 117,   0, 111, 218, 228, 226, 192, 152, 169, 180, 159, 126, 251, 117, 211,  48, 135, 121, 229,   0)
        m_gp(26) = Array( 70, 218, 145, 153, 227,  48, 102,  13, 142, 245,  21, 161,  53, 165,  28, 111, 201, 145,  17, 118, 182, 103,   2, 158, 125, 173,   0)
        m_gp(28) = Array(123,   9,  37, 242, 119, 212, 195,  42,  87, 245,  43,  21, 201, 232,  27, 205, 147, 195, 190, 110, 180, 108, 234, 224, 104, 200, 223, 168,   0)
        m_gp(30) = Array(180, 192,  40, 238, 216, 251,  37, 156, 130, 224, 193, 226, 173,  42, 125, 222,  96, 239,  86, 110,  48,  50, 182, 179,  31, 216, 152, 145, 173,  41,   0)
        m_gp(32) = Array(241, 220, 185, 254,  52,  80, 222,  28,  60, 171,  60,  38, 156,  80, 185, 120,  27,  89, 123, 242,  32, 138, 138, 209,  67,   4, 167, 249, 190, 106,   6,  10,   0)
        m_gp(34) = Array( 51, 129,  62,  98,  13, 167, 129, 183,  61, 114,  70,  56, 103, 218, 239, 229, 158,  58, 125, 163, 140,  86, 193, 113,  94, 105,  19, 108,  21,  26,  94, 146,  77, 111,   0)
        m_gp(36) = Array(120,  30, 233, 113, 251, 117, 196, 121,  74, 120, 177, 105, 210,  87,  37, 218,  63,  18, 107, 238, 248, 113, 152, 167,   0, 115, 152,  60, 234, 246,  31, 172,  16,  98, 183, 200,   0)
        m_gp(40) = Array( 15,  35,  53, 232,  20,  72, 134, 125, 163,  47,  41,  88, 114, 181,  35, 175,   7, 170, 104, 226, 174, 187,  26,  53, 106, 235,  56, 163,  57, 247, 161, 128, 205, 128,  98, 252, 161,  79, 116,  59,   0)
        m_gp(42) = Array( 96,  50, 117, 194, 162, 171, 123, 201, 254, 237, 199, 213, 101,  39, 223, 101,  34, 139, 131,  15, 147,  96, 106, 188,   8, 230,  84, 110, 191, 221, 242,  58,   3,   0, 231, 137,  18,  25, 230, 221, 103, 250,   0)
        m_gp(44) = Array(181,  73, 102, 113, 130,  37, 169, 204, 147, 217, 194,  52, 163,  68, 114, 118, 126, 224,  62, 143,  78,  44, 238,   1, 247,  14, 145,   9, 123,  72,  25, 191, 243,  89, 188, 168,  55,  69, 246,  71, 121,  61,   7, 190,   0)
        m_gp(46) = Array( 15,  82,  19, 223, 202,  43, 224, 157,  25,  52, 174, 119, 245, 249,   8, 234, 104,  73, 241,  60,  96,   4,   1,  36, 211, 169, 216, 135,  16,  58,  44, 129, 113,  54,   5,  89,  99, 187, 115, 202, 224, 253, 112,  88,  94, 112,   0)
        m_gp(48) = Array(108,  34,  39, 163,  50,  84, 227,  94,  11, 191, 238, 140, 156, 247,  21,  91, 184, 120, 150,  95, 206, 107, 205, 182, 160, 135, 111, 221,  18, 115, 123,  46,  63, 178,  61, 240, 102,  39,  90, 251,  24,  60, 146, 211, 130, 196,  25, 228,   0)
        m_gp(50) = Array(205, 133, 232, 215, 170, 124, 175, 235, 114, 228,  69, 124,  65, 113,  32, 189,  42,  77,  75, 242, 215, 242, 160, 130, 209, 126, 160,  32,  13,  46, 225, 203, 242, 195, 111, 209,   3,  35, 193, 203,  99, 209,  46, 118,   9, 164, 161, 157, 125, 232,   0)
        m_gp(52) = Array( 51, 116, 254, 239,  33, 101, 220, 200, 242,  39,  97,  86,  76,  22, 121, 235, 233, 100, 113, 124,  65,  59,  94, 190,  89, 254, 134, 203, 242,  37, 145,  59,  14,  22, 215, 151, 233, 184,  19, 124, 127,  86,  46, 192,  89, 251, 220,  50, 186,  86,  50, 116,   0)
        m_gp(54) = Array(156,  31,  76, 198,  31, 101,  59, 153,   8, 235, 201, 128,  80, 215, 108, 120,  43, 122,  25, 123,  79, 172, 175, 238, 254,  35, 245,  52, 192, 184,  95,  26, 165, 109, 218, 209,  58, 102, 225, 249, 184, 238,  50,  45,  65,  46,  21, 113, 221, 210,  87, 201,  26, 183,   0)
        m_gp(56) = Array( 10,  61,  20, 207, 202, 154, 151, 247, 196,  27,  61, 163,  23,  96, 206, 152, 124, 101, 184, 239,  85,  10,  28, 190, 174, 177, 249, 182, 142, 127, 139,  12, 209, 170, 208, 135, 155, 254, 144,   6, 229, 202, 201,  36, 163, 248,  91,   2, 116, 112, 216, 164, 157, 107, 120, 106,   0)
        m_gp(58) = Array(123, 148, 125, 233, 142, 159,  63,  41,  29, 117, 245, 206, 134, 127, 145,  29, 218, 129,   6, 214, 240, 122,  30,  24,  23, 125, 165,  65, 142, 253,  85, 206, 249, 152, 248, 192, 141, 176, 237, 154, 144, 210, 242, 251,  55, 235, 185, 200, 182, 252, 107,  62,  27,  66, 247,  26, 116,  82,   0)
        m_gp(60) = Array(240,  33,   7,  89,  16, 209,  27,  70, 220, 190, 102,  65,  87, 194,  25,  84, 181,  30, 124,  11,  86, 121, 209, 160,  49, 238,  38,  37,  82, 160, 109, 101, 219, 115,  57, 198, 205,   2, 247, 100,   6, 127, 181,  28, 120, 219, 101, 211,  45, 219, 197, 226, 197, 243, 141,   9,  12,  26, 140, 107,   0)
        m_gp(62) = Array(106, 110, 186,  36, 215, 127, 218, 182, 246,  26, 100, 200,   6, 115,  40, 213, 123, 147, 149, 229,  11, 235, 117, 221,  35, 181, 126, 212,  17, 194, 111,  70,  50,  72,  89, 223,  76,  70, 118, 243,  78, 135, 105,   7, 121,  58, 228,   2,  23,  37, 122,   0,  94, 214, 118, 248, 223,  71,  98, 113, 202,  65,   0)
        m_gp(64) = Array(231, 213, 156, 217, 243, 178,  11, 204,  31, 242, 230, 140, 108,  99,  63, 238, 242, 125, 195, 195, 140,  47, 146, 184,  47,  91, 216,   4, 209, 218, 150, 208, 156, 145,  24,  29, 212, 199,  93, 160,  53, 127,  26, 119, 149, 141,  78, 200, 254, 187, 204, 177, 123,  92, 119,  68,  49, 159, 158,   7,   9, 175,  51,  45,   0)
        m_gp(66) = Array(105,  45,  93, 132,  25, 171, 106,  67, 146,  76,  82, 168,  50, 106, 232,  34,  77, 217, 126, 240, 253,  80,  87,  63, 143, 121,  40, 236, 111,  77, 154,  44,   7,  95, 197, 169, 214,  72,  41, 101,  95, 111,  68, 178, 137,  65, 173,  95, 171, 197, 247, 139,  17,  81, 215,  13, 117,  46,  51, 162, 136, 136, 180, 222, 118,   5,   0)
        m_gp(68) = Array(238, 163,   8,   5,   3, 127, 184, 101,  27, 235, 238,  43, 198, 175, 215,  82,  32,  54,   2, 118, 225, 166, 241, 137, 125,  41, 177,  52, 231,  95,  97, 199,  52, 227,  89, 160, 173, 253,  84,  15,  84,  93, 151, 203, 220, 165, 202,  60,  52, 133, 205, 190, 101,  84, 150,  43, 254,  32, 160,  90,  70,  77,  93, 224,  33, 223, 159, 247,   0)
    End Sub

    Public Function Item(ByVal numECCodewords)
        If IsEmpty(m_gp(numECCodewords)) Then Call Err.Raise(5)

        Item = m_gp(numECCodewords)
    End Function
End Class


Class KanjiEncoder
    Private m_data
    Private m_charCounter
    Private m_bitCounter

    Private m_encAlpha

    Private Sub Class_Initialize()
        m_data = Empty
        m_charCounter = 0
        m_bitCounter = 0

        Set m_encAlpha = New AlphanumericEncoder
    End Sub

    Public Property Get BitCount()
        BitCount = m_bitCounter
    End Property

    Public Property Get CharCount()
        CharCount = m_charCounter
    End Property

    Public Property Get EncodingMode()
        EncodingMode = MODE_KANJI
    End Property

    Public Property Get ModeIndicator()
        ModeIndicator = MODEINDICATOR_KANJI_VALUE
    End Property

    Public Function Append(ByVal c)
        Dim wd
        wd = Asc(c) And &HFFFF&

        If &H8140& <= wd And wd <= &H9FFC& Then
            wd = wd - &H8140&
        ElseIf &HE040& <= wd And wd <= &HEBBF& Then
            wd = wd - &HC140&
        Else
            Call Err.Raise(5)
        End If

        wd = ((wd \ 2 ^ 8) * &HC0&) + (wd And &HFF&)
        If m_charCounter = 0 Then
            ReDim m_data(0)
        Else
            ReDim Preserve m_data(UBound(m_data) + 1)
        End If

        m_data(UBound(m_data)) = wd

        Dim ret
        ret = GetCodewordBitLength(c)
        m_bitCounter = m_bitCounter + ret
        m_charCounter = m_charCounter + 1

        Append = ret
    End Function

    Public Function GetCodewordBitLength(ByVal c)
        GetCodewordBitLength = 13
    End Function

    Public Function GetBytes()
        Dim bs
        Set bs = New BitSequence

        Dim v
        For Each v In m_data
            Call bs.Append(v, 13)
        Next

        GetBytes = bs.GetBytes()
    End Function

    Public Function InSubset(ByVal c)
        Dim code
        code = Asc(c) And &HFFFF&

        Dim lsb
        lsb = code And &HFF&

        If &H8140& <= code And code <= &H9FFC& Or _
           &HE040& <= code And code <= &HEBBF& Then
            InSubset = &H40& <= lsb And lsb <= &HFC& And _
                       &H7F& <> lsb
            Exit Function
        End If

        InSubset = False
    End Function

    Public Function InExclusiveSubset(ByVal c)
        If m_encAlpha.InSubset(c) Then
            InExclusiveSubset = False
        End If

        InExclusiveSubset = InSubset(c)
    End Function
End Class


Class List
    Private m_items

    Private Sub Class_Initialize()
        m_items = Array()
    End Sub

    Public Sub Add(arg)
        ReDim Preserve m_items(UBound(m_items) + 1)

        If VarType(arg) = vbObject Then
            Set m_items(UBound(m_items)) = arg
        Else
            m_items(UBound(m_items)) = arg
        End If
    End Sub

    Public Property Get Count()
        Count = UBound(m_items) + 1
    End Property

    Public Property Get Item(ByVal idx)
        If VarType(m_items(idx)) = vbObject Then
            Set Item = m_items(idx)
        Else
            Item = m_items(idx)
        End If
    End Property

    Public Property Get Items()
        Items = m_items
    End Property
End Class


Class Masking_
    Public Function Apply(ByVal ver, ByVal ecLevel, ByRef moduleMatrix)
        Dim minPenalty
        minPenalty = &H7FFFFFFF

        Dim temp
        Dim penalty
        Dim maskPatternReference
        Dim maskedMatrix

        Dim i

        For i = 0 To 7
            temp = moduleMatrix

            Call Mask(i, temp)
            Call FormatInfo.Place(ecLevel, i, temp)

            If ver >= 7 Then
                Call VersionInfo.Place(ver, temp)
            End If

            penalty = MaskingPenaltyScore.CalcTotal(temp)

            If penalty < minPenalty Then
                minPenalty = penalty
                maskPatternReference = i
                maskedMatrix = temp
            End If
        Next

        moduleMatrix = maskedMatrix
        Apply = maskPatternReference
    End Function

    Private Sub Mask(ByVal maskPatternReference, ByRef moduleMatrix())
        Dim condition
        Set condition = GetCondition(maskPatternReference)

        Dim r, c

        For r = 0 To UBound(moduleMatrix)
            For c = 0 To UBound(moduleMatrix(r))
                If Abs(moduleMatrix(r)(c)) = WORD Then
                    If condition.Evaluate(r, c) Then
                        moduleMatrix(r)(c) = moduleMatrix(r)(c) * -1
                    End If
                End If
            Next
        Next
    End Sub

    Private Function GetCondition(ByVal maskPatternReference)
        Dim ret

        Select Case maskPatternReference
            Case 0
                Set ret = New MaskingCondition0
            Case 1
                Set ret = New MaskingCondition1
            Case 2
                Set ret = New MaskingCondition2
            Case 3
                Set ret = New MaskingCondition3
            Case 4
                Set ret = New MaskingCondition4
            Case 5
                Set ret = New MaskingCondition5
            Case 6
                Set ret = New MaskingCondition6
            Case 7
                Set ret = New MaskingCondition7
            Case Else
                Call Err.Raise(5)
        End Select

        Set GetCondition = ret
    End Function
End Class


Class MaskingCondition0
    Public Function Evaluate(ByVal r, ByVal c)
        Evaluate = (r + c) Mod 2 = 0
    End Function
End Class


Class MaskingCondition1
    Public Function Evaluate(ByVal r, ByVal c)
        Evaluate = r Mod 2 = 0
    End Function
End Class


Class MaskingCondition2
    Public Function Evaluate(ByVal r, ByVal c)
        Evaluate = c Mod 3 = 0
    End Function
End Class


Class MaskingCondition3
    Public Function Evaluate(ByVal r, ByVal c)
        Evaluate = (r + c) Mod 3 = 0
    End Function
End Class


Class MaskingCondition4
    Public Function Evaluate(ByVal r, ByVal c)
        Evaluate = ((r \ 2) + (c \ 3)) Mod 2 = 0
    End Function
End Class


Class MaskingCondition5
    Public Function Evaluate(ByVal r, ByVal c)
        Evaluate = ((r * c) Mod 2 + (r * c) Mod 3) = 0
    End Function
End Class


Class MaskingCondition6
    Public Function Evaluate(ByVal r, ByVal c)
        Evaluate = ((r * c) Mod 2 + (r * c) Mod 3) Mod 2 = 0
    End Function
End Class


Class MaskingCondition7
    Public Function Evaluate(ByVal r, ByVal c)
        Evaluate = ((r + c) Mod 2 + (r * c) Mod 3) Mod 2 = 0
    End Function
End Class


Class MaskingPenaltyScore_
    Public Function CalcTotal(ByRef moduleMatrix())
        Dim total
        Dim penalty

        penalty = CalcAdjacentModulesInSameColor(moduleMatrix)
        total = total + penalty

        penalty = CalcBlockOfModulesInSameColor(moduleMatrix)
        total = total + penalty

        penalty = CalcModuleRatio(moduleMatrix)
        total = total + penalty

        penalty = CalcProportionOfDarkModules(moduleMatrix)
        total = total + penalty

        CalcTotal = total
    End Function

    Private Function CalcAdjacentModulesInSameColor(ByRef moduleMatrix())
        Dim penalty
        penalty = 0

        penalty = penalty + CalcAdjacentModulesInRowInSameColor(moduleMatrix)
        penalty = penalty + CalcAdjacentModulesInRowInSameColor(MatrixRotate90(moduleMatrix))

        CalcAdjacentModulesInSameColor = penalty
    End Function

    Private Function CalcAdjacentModulesInRowInSameColor(ByRef moduleMatrix())
        Dim penalty
        penalty = 0

        Dim rowArray
        Dim i
        Dim cnt

        For Each rowArray In moduleMatrix
            cnt = 1

            For i = 0 To UBound(rowArray) - 1
                If IsDark(rowArray(i)) = IsDark(rowArray(i + 1)) Then
                    cnt = cnt + 1
                Else
                    If cnt >= 5 Then
                        penalty = penalty + (3 + (cnt - 5))
                    End If

                    cnt = 1
                End If
            Next

            If cnt >= 5 Then
                penalty = penalty + (3 + (cnt - 5))
            End If
        Next

        CalcAdjacentModulesInRowInSameColor = penalty
    End Function

    Private Function CalcBlockOfModulesInSameColor(ByRef moduleMatrix())
        Dim penalty
        Dim r, c
        Dim temp

        For r = 0 To UBound(moduleMatrix) - 1
            For c = 0 To UBound(moduleMatrix(r)) - 1
                temp = IsDark(moduleMatrix(r)(c))

                If (IsDark(moduleMatrix(r + 0)(c + 1)) = temp) And _
                   (IsDark(moduleMatrix(r + 1)(c + 0)) = temp) And _
                   (IsDark(moduleMatrix(r + 1)(c + 1)) = temp) Then
                    penalty = penalty + 3
                End If
            Next
        Next

        CalcBlockOfModulesInSameColor = penalty
    End Function

    Private Function CalcModuleRatio(ByRef moduleMatrix())
        Dim moduleMatrixTemp
        moduleMatrixTemp = QuietZone.Place(moduleMatrix)

        Dim penalty
        penalty = 0

        penalty = penalty + CalcModuleRatioInRow(moduleMatrixTemp)
        penalty = penalty + CalcModuleRatioInRow(MatrixRotate90(moduleMatrixTemp))

        CalcModuleRatio = penalty
    End Function

    Private Function CalcModuleRatioInRow(ByRef moduleMatrix())
        Dim penalty

        Dim ratio3Ranges
        Dim rowArray

        Dim ratio1, ratio3, ratio4

        Dim i
        Dim cnt
        Dim flg
        Dim impose

        Dim rng

        For Each rowArray In moduleMatrix
            ratio3Ranges = GetRatio3Ranges(rowArray)

            For Each rng In ratio3Ranges
                ratio3 = rng(1) + 1 - rng(0)
                ratio1 = ratio3 \ 3
                ratio4 = ratio1 * 4
                flg = True
                impose = False

                i = rng(0) - 1

                If flg Then
                    ' light ratio 1
                    cnt = 0
                    Do While i >= 0
                        If Not IsDark(rowArray(i)) Then
                            cnt = cnt + 1
                            i = i - 1
                        Else
                            Exit Do
                        End If
                    Loop

                    flg = cnt = ratio1
                End If

                If flg Then
                    ' dark ratio 1
                    cnt = 0
                    Do While i >= 0
                        If IsDark(rowArray(i)) Then
                            cnt = cnt + 1
                            i = i - 1
                        Else
                            Exit Do
                        End If
                    Loop

                    flg = cnt = ratio1
                End If

                If flg Then
                    ' light ratio 4
                    cnt = 0
                    Do While i >= 0
                        If Not IsDark(rowArray(i)) Then
                            cnt = cnt + 1
                            i = i - 1
                        Else
                            Exit Do
                        End If
                    Loop

                    If cnt >= ratio4 Then
                        impose = True
                    End If
                End If

                i = rng(1) + 1

                If flg Then
                    ' light ratio 1
                    cnt = 0
                    Do While i <= UBound(rowArray)
                        If Not IsDark(rowArray(i)) Then
                            cnt = cnt + 1
                            i = i + 1
                        Else
                            Exit Do
                        End If
                    Loop

                    flg = cnt = ratio1
                End If

                If flg Then
                    ' dark ratio 1
                    cnt = 0
                    Do While i <= UBound(rowArray)
                        If IsDark(rowArray(i)) Then
                            cnt = cnt + 1
                            i = i + 1
                        Else
                            Exit Do
                        End If
                    Loop

                    flg = cnt = ratio1
                End If

                If flg Then
                    ' light ratio 4
                    cnt = 0
                    Do While i <= UBound(rowArray)
                        If Not IsDark(rowArray(i)) Then
                            cnt = cnt + 1
                            i = i + 1
                        Else
                            Exit Do
                        End If
                    Loop

                    If cnt >= ratio4 Then
                        impose = True
                    End If
                End If

                If flg And impose Then
                    penalty = penalty + 40
                End If
            Next
        Next

        CalcModuleRatioInRow = penalty
    End Function

    Private Function GetRatio3Ranges(ByRef arg)
        Dim ret
        ret = Array()

        Dim s, i

        For i = 1 To UBound(arg) - 1
            If IsDark(arg(i)) Then
                If Not IsDark(arg(i - 1)) Then
                    s = i
                End If

                If Not IsDark(arg(i + 1)) Then
                    If (i + 1 - s) Mod 3 = 0 Then
                        ReDim Preserve ret(UBound(ret) + 1)
                        ret(UBound(ret)) = Array(s, i)
                    End If
                End If
            End If
        Next

        GetRatio3Ranges = ret
    End Function

    Private Function CalcProportionOfDarkModules(ByRef moduleMatrix())
        Dim darkCount

        Dim rowArray
        Dim v

        For Each rowArray In moduleMatrix
            For Each v In rowArray
                If IsDark(v) Then
                    darkCount = darkCount + 1
                End If
            Next
        Next

        Dim numModules
        numModules = (UBound(moduleMatrix) + 1) ^ 2

        Dim k
        k = darkCount / numModules * 100
        k = Abs(k - 50)
        k = Int(k / 5)
        Dim penalty
        penalty = CInt(k) * 10

        CalcProportionOfDarkModules = penalty
    End Function

    Private Function MatrixRotate90(ByRef arg())
        Dim ret()
        ReDim ret(UBound(arg(0)))

        Dim i, j
        Dim cols()

        For i = 0 To UBound(ret)
            ReDim cols(UBound(arg))
            ret(i) = cols
        Next

        Dim k
        k = UBound(ret)

        For i = 0 To UBound(ret)
            For j = 0 To UBound(ret(i))
                ret(i)(j) = arg(j)(k - i)
            Next
        Next

        MatrixRotate90 = ret
    End Function
End Class


Class Module_
    Public Function GetNumModulesPerSide(ByVal ver)
        GetNumModulesPerSide = 17 + ver * 4
    End Function
End Class


Class NumericEncoder
    Private m_data
    Private m_charCounter
    Private m_bitCounter

    Private Sub Class_Initialize()
        m_data = Empty
        m_charCounter = 0
        m_bitCounter = 0
    End Sub

    Public Property Get BitCount()
        BitCount = m_bitCounter
    End Property

    Public Property Get CharCount()
        CharCount = m_charCounter
    End Property

    Public Property Get EncodingMode()
        EncodingMode = MODE_NUMERIC
    End Property

    Public Property Get ModeIndicator()
        ModeIndicator = MODEINDICATOR_NUMERIC_VALUE
    End Property

    Public Function Append(ByVal c)
        If m_charCounter Mod 3 = 0 Then
            If m_charCounter = 0 Then
                ReDim m_data(0)
            Else
                ReDim Preserve m_data(UBound(m_data) + 1)
            End If

            m_data(UBound(m_data)) = CLng(c)
        Else
            m_data(UBound(m_data)) = m_data(UBound(m_data)) * 10 + CLng(c)
        End If

        Dim ret
        ret = GetCodewordBitLength(c)
        m_bitCounter = m_bitCounter + ret
        m_charCounter = m_charCounter + 1

        Append = ret
    End Function

    Public Function GetCodewordBitLength(ByVal c)
        If m_charCounter Mod 3 = 0 Then
            GetCodewordBitLength = 4
        Else
            GetCodewordBitLength = 3
        End If
    End Function

    Public Function GetBytes()
        Dim bs
        Set bs = New BitSequence

        Dim i
        For i = 0 To UBound(m_data) - 1
            Call bs.Append(m_data(i), 10)
        Next

        Select Case m_charCounter Mod 3
            Case 1
                Call bs.Append(m_data(UBound(m_data)), 4)
            Case 2
                Call bs.Append(m_data(UBound(m_data)), 7)
            Case Else
                Call bs.Append(m_data(UBound(m_data)), 10)
        End Select

        GetBytes = bs.GetBytes()
    End Function

    Public Function InSubset(ByVal c)
        InSubset = "0" <= c And c <= "9"
    End Function

    Public Function InExclusiveSubset(ByVal c)
        InExclusiveSubset = InSubset(c)
    End Function
End Class


Class QuietZone_
    Private m_width

    Private Sub Class_Initialize()
        m_width = QUIET_ZONE_MIN_WIDTH
    End Sub

    Public Property Get Width()
        Width = m_width
    End Property
    Public Property Let Width(ByVal Value)
        If Value < QUIET_ZONE_MIN_WIDTH Then
            Call Err.Raise(5)
        End If

        m_width = Value
    End Property

    Public Function Place(ByRef moduleMatrix())
        Dim ret()
        ReDim ret(UBound(moduleMatrix) + Width * 2)

        Dim i
        Dim cols()

        For i = 0 To UBound(ret)
            ReDim cols(UBound(ret))
            ret(i) = cols
        Next

        Dim r, c

        For r = 0 To UBound(moduleMatrix)
            For c = 0 To UBound(moduleMatrix(r))
                ret(r + Width)(c + Width) = moduleMatrix(r)(c)
            Next
        Next

        Place = ret
    End Function
End Class


Class RemainderBit_
    Public Sub Place(ByRef moduleMatrix())
        Dim r, c

        For r = 0 To UBound(moduleMatrix)
            For c = 0 To UBound(moduleMatrix(r))
                If moduleMatrix(r)(c) = BLANK Then
                    moduleMatrix(r)(c) = -WORD
                End If
            Next
        Next
    End Sub
End Class


Class RSBlock_
    Private m_totalNumbers

    Private Sub Class_Initialize()
        m_totalNumbers = Array( _
            Array( 0, _
                   1,  1,  1,  1,  1,  2,  2,  2,  2,  4, _
                   4,  4,  4,  4,  6,  6,  6,  6,  7,  8, _
                   8,  9,  9, 10, 12, 12, 12, 13, 14, 15, _
                  16, 17, 18, 19, 19, 20, 21, 22, 24, 25), _
            Array( 0, _
                   1,  1,  1,  2,  2,  4,  4,  4,  5,  5, _
                   5,  8,  9,  9, 10, 10, 11, 13, 14, 16, _
                  17, 17, 18, 20, 21, 23, 25, 26, 28, 29, _
                  31, 33, 35, 37, 38, 40, 43, 45, 47, 49), _
            Array( 0, _
                   1,  1,  2,  2,  4,  4,  6,  6,  8,  8, _
                   8, 10, 12, 16, 12, 17, 16, 18, 21, 20, _
                  23, 23, 25, 27, 29, 34, 34, 35, 38, 40, _
                  43, 45, 48, 51, 53, 56, 59, 62, 65, 68), _
            Array( 0, _
                   1,  1,  2,  4,  4,  4,  5,  6,  8,  8, _
                  11, 11, 16, 16, 18, 16, 19, 21, 25, 25, _
                  25, 34, 30, 32, 35, 37, 40, 42, 45, 48, _
                  51, 54, 57, 60, 63, 66, 70, 74, 77, 81) _
        )
    End Sub

    Public Function GetTotalNumber(ByVal ecLevel, ByVal ver, ByVal preceding)
        Dim dataWordCapacity
        Dim blockCount

        dataWordCapacity = DataCodeword.GetTotalNumber(ecLevel, ver)
        blockCount = m_totalNumbers(ecLevel)

        If preceding Then
            GetTotalNumber = blockCount(ver) - (dataWordCapacity Mod blockCount(ver))
        Else
            GetTotalNumber = dataWordCapacity Mod blockCount(ver)
        End If
    End Function

    Public Function GetNumberDataCodewords(ByVal ecLevel, ByVal ver, ByVal preceding)
        Dim ret

        Dim numDataCodewords
        numDataCodewords = DataCodeword.GetTotalNumber(ecLevel, ver)

        Dim numBlocks
        numBlocks = m_totalNumbers(ecLevel)(ver)

        Dim numPreBlockCodewords
        numPreBlockCodewords = numDataCodewords \ numBlocks

        Dim numPreBlocks
        Dim numFolBlocks

        If preceding Then
            ret = numPreBlockCodewords
        Else
            numPreBlocks = GetTotalNumber(ecLevel, ver, True)
            numFolBlocks = GetTotalNumber(ecLevel, ver, False)

            If numFolBlocks > 0 Then
                ret = (numDataCodewords - numPreBlockCodewords * numPreBlocks) \ numFolBlocks
            Else
                ret = 0
            End If
        End If

        GetNumberDataCodewords = ret
    End Function

    Public Function GetNumberECCodewords(ByVal ecLevel, ByVal ver)
        Dim numDataCodewords
        numDataCodewords = DataCodeword.GetTotalNumber(ecLevel, ver)

        Dim numBlocks
        numBlocks = m_totalNumbers(ecLevel)(ver)

        GetNumberECCodewords = _
            (Codeword.GetTotalNumber(ver) \ numBlocks) - _
                (numDataCodewords \ numBlocks)
    End Function
End Class


Class Separator_
    Public Sub Place(ByRef moduleMatrix())
        DIM VAL
        VAL = SEPARATOR_PTN

        Dim offset
        offset = UBound(moduleMatrix) - 7

        Dim i
        For i = 0 To 7
             moduleMatrix(i)(7) = -VAL
             moduleMatrix(7)(i) = -VAL

             moduleMatrix(offset + i)(7) = -VAL
             moduleMatrix(offset + 0)(i) = -VAL

             moduleMatrix(i)(offset + 0) = -VAL
             moduleMatrix(7)(offset + i) = -VAL
         Next
    End Sub
End Class


Class Symbol
    Private m_parent

    Private m_position

    Private m_currEncoder
    Private m_currEncodingMode
    Private m_currVersion

    Private m_dataBitCapacity
    Private m_dataBitCounter

    Private m_segments
    Private m_segmentCounter

    Private Sub Class_Initialize()
        Set m_segments = New List
        Set m_segmentCounter = CreateObject("Scripting.Dictionary")
    End Sub

    Public Sub Init(ByVal parentObj)
        Set m_parent = parentObj

        m_position = parentObj.Count

        Set m_currEncoder = Nothing
        m_currEncodingMode = MODE_UNKNOWN
        m_currVersion = parentObj.MinVersion

        m_dataBitCapacity = 8 * DataCodeword.GetTotalNumber( _
            parentObj.ErrorCorrectionLevel, parentObj.MinVersion)

        m_dataBitCounter = 0

        Call m_segmentCounter.Add(MODE_NUMERIC, 0)
        Call m_segmentCounter.Add(MODE_ALPHA_NUMERIC, 0)
        Call m_segmentCounter.Add(MODE_BYTE, 0)
        Call m_segmentCounter.Add(MODE_KANJI, 0)

        If parentObj.StructuredAppendAllowed Then
            m_dataBitCapacity = m_dataBitCapacity - STRUCTUREDAPPEND_HEADER_LENGTH
        End If
    End Sub

    Public Property Get Parent()
        Set Parent = m_parent
    End Property

    Public Property Get Version()
        Version = m_currVersion
    End Property

    Public Property Get CurrentEncodingMode()
        CurrentEncodingMode = m_currEncodingMode
    End Property

    Public Function TryAppend(ByVal c)
        Dim bitLength
        bitLength = m_currEncoder.GetCodewordBitLength(c)

        Do While (m_dataBitCapacity < m_dataBitCounter + bitLength)
            If m_currVersion >= m_parent.MaxVersion Then
                TryAppend = False
                Exit Function
            End If

            Call SelectVersion
        Loop

        Call m_currEncoder.Append(c)
        m_dataBitCounter = m_dataBitCounter + bitLength
        Call m_parent.UpdateParity(c)

        TryAppend = True
    End Function

    Public Function TrySetEncodingMode(ByVal encMode, ByVal c)
        Dim encoder
        Set encoder = CreateEncoder(encMode)

        Dim bitLength
        bitLength = encoder.GetCodewordBitLength(c)

        Do While (m_dataBitCapacity < _
                    m_dataBitCounter + _
                    MODEINDICATOR_LENGTH + _
                    CharCountIndicator.GetLength(m_currVersion, encMode) + _
                    bitLength)

            If m_currVersion >= m_parent.MaxVersion Then
                TrySetEncodingMode = False
                Exit Function
            End If

            Call SelectVersion
        Loop

        m_dataBitCounter = m_dataBitCounter + _
                           MODEINDICATOR_LENGTH + _
                           CharCountIndicator.GetLength(m_currVersion, encMode)

        Set m_currEncoder = encoder
        Call m_segments.Add(encoder)
        m_segmentCounter(encMode) = m_segmentCounter(encMode) + 1
        m_currEncodingMode = encMode

        TrySetEncodingMode = True
    End Function

    Private Sub SelectVersion()
        Dim encMode
        Dim num

        For Each encMode In m_segmentCounter.Keys()
            num = m_segmentCounter(encMode)

            m_dataBitCounter = m_dataBitCounter + _
                               num * CharCountIndicator.GetLength( _
                                    m_currVersion + 1, encMode) - _
                               num * CharCountIndicator.GetLength( _
                                    m_currVersion + 0, encMode)
        Next

        m_currVersion = m_currVersion + 1
        m_dataBitCapacity = 8 * DataCodeword.GetTotalNumber( _
            m_parent.ErrorCorrectionLevel, m_currVersion)
        m_parent.MinVersion = m_currVersion

        If m_parent.StructuredAppendAllowed Then
            m_dataBitCapacity = m_dataBitCapacity - STRUCTUREDAPPEND_HEADER_LENGTH
        End If
    End Sub

    Private Function BuildDataBlock()
        Dim dataBytes
        dataBytes = GetMessageBytes()

        Dim numPreBlocks
        numPreBlocks = RSBlock.GetTotalNumber( _
            m_parent.ErrorCorrectionLevel, m_currVersion, True)

        Dim numFolBlocks
        numFolBlocks = RSBlock.GetTotalNumber( _
            m_parent.ErrorCorrectionLevel, m_currVersion, False)

        Dim ret()
        ReDim ret(numPreBlocks + numFolBlocks - 1)

        Dim dataIdx
        dataIdx = 0

        Dim numPreBlockDataCodewords
        numPreBlockDataCodewords = RSBlock.GetNumberDataCodewords( _
            m_parent.ErrorCorrectionLevel, m_currVersion, True)

        Dim data()
        Dim i, j

        For i = 0 To numPreBlocks - 1
            ReDim data(numPreBlockDataCodewords - 1)

            For j = 0 To UBound(data)
                data(j) = dataBytes(dataIdx)
                dataIdx = dataIdx + 1
            Next

            ret(i) = data
        Next

        Dim numFolBlockDataCodewords
        numFolBlockDataCodewords = RSBlock.GetNumberDataCodewords( _
            m_parent.ErrorCorrectionLevel, m_currVersion, False)

        For i = numPreBlocks To numPreBlocks + numFolBlocks - 1
            ReDim data(numFolBlockDataCodewords - 1)

            For j = 0 To UBound(data)
                data(j) = dataBytes(dataIdx)
                dataIdx = dataIdx + 1
            Next

            ret(i) = data
        Next

        BuildDataBlock = ret
    End Function

    Private Function BuildErrorCorrectionBlock(ByRef dataBlock())
        Dim i, j

        Dim numECCodewords
        numECCodewords = RSBlock.GetNumberECCodewords( _
            m_parent.ErrorCorrectionLevel, m_currVersion)

        Dim numPreBlocks
        numPreBlocks = RSBlock.GetTotalNumber( _
            m_parent.ErrorCorrectionLevel, m_currVersion, True)

        Dim numFolBlocks
        numFolBlocks = RSBlock.GetTotalNumber( _
            m_parent.ErrorCorrectionLevel, m_currVersion, False)

        Dim ret()
        ReDim ret(numPreBlocks + numFolBlocks - 1)

        Dim eccDataTmp()
        ReDim eccDataTmp(numECCodewords - 1)

        For i = 0 To UBound(ret)
            ret(i) = eccDataTmp
        Next

        Dim gp
        gp = GeneratorPolynomials.Item(numECCodewords)

        Dim eccIdx
        Dim blockIdx
        Dim data()
        Dim exp

        For blockIdx = 0 To UBound(dataBlock)
            ReDim data(UBound(dataBlock(blockIdx)) + UBound(ret(blockIdx)) + 1)
            eccIdx = UBound(data)

            For i = 0 To UBound(dataBlock(blockIdx))
                data(eccIdx) = dataBlock(blockIdx)(i)
                eccIdx = eccIdx - 1
            Next

            For i = UBound(data) To numECCodewords Step -1
                If data(i) > 0 Then
                    exp = GaloisField256.ToExp(data(i))
                    eccIdx = i

                    For j = UBound(gp) To 0 Step -1
                        data(eccIdx) = data(eccIdx) Xor _
                                       GaloisField256.ToInt((gp(j) + exp) Mod 255)
                        eccIdx = eccIdx - 1
                    Next
                End If
            Next

            eccIdx = numECCodewords - 1

            For i = 0 To UBound(ret(blockIdx))
                ret(blockIdx)(i) = data(eccIdx)
                eccIdx = eccIdx - 1
            Next
        Next

        BuildErrorCorrectionBlock = ret
    End Function

    Private Function GetEncodingRegionBytes()
        Dim dataBlock
        dataBlock = BuildDataBlock()

        Dim ecBlock
        ecBlock = BuildErrorCorrectionBlock(dataBlock)

        Dim numCodewords
        numCodewords = Codeword.GetTotalNumber(m_currVersion)

        Dim numDataCodewords
        numDataCodewords = DataCodeword.GetTotalNumber( _
            m_parent.ErrorCorrectionLevel, m_currVersion)

        Dim ret()
        ReDim ret(numCodewords - 1)

        Dim r, c

        Dim idx
        idx = 0

        Dim n
        n = 0

        Do While idx < numDataCodewords
            r = n Mod (UBound(dataBlock) + 1)
            c = n \ (UBound(dataBlock) + 1)

            If c <= UBound(dataBlock(r)) Then
                ret(idx) = dataBlock(r)(c)
                idx = idx + 1
            End If

            n = n + 1
        Loop

        n = 0

        Do While idx < numCodewords
            r = n Mod (UBound(ecBlock) + 1)
            c = n \ (UBound(ecBlock) + 1)

            If c <= UBound(ecBlock(r)) Then
                ret(idx) = ecBlock(r)(c)
                idx = idx + 1
            End If

            n = n + 1
        Loop

        GetEncodingRegionBytes = ret
    End Function

    Private Function GetMessageBytes()
        Dim bs
        Set bs = New BitSequence

        If m_parent.Count > 1 Then
            Call WriteStructuredAppendHeader(bs)
        End If

        Call WriteSegments(bs)
        Call WriteTerminator(bs)
        Call WritePaddingBits(bs)
        Call WritePadCodewords(bs)

        GetMessageBytes = bs.GetBytes()
    End Function

    Private Sub WriteStructuredAppendHeader(ByVal bs)
        Call bs.Append(MODEINDICATOR_STRUCTURED_APPEND_VALUE, _
                       MODEINDICATOR_LENGTH)
        Call bs.Append(m_position, _
                       SYMBOLSEQUENCEINDICATOR_POSITION_LENGTH)
        Call bs.Append(m_parent.Count - 1, _
                       SYMBOLSEQUENCEINDICATOR_TOTAL_NUMBER_LENGTH)
        Call bs.Append(m_parent.Parity, _
                       STRUCTUREDAPPEND_PARITY_DATA_LENGTH)
    End Sub

    Private Sub WriteSegments(ByVal bs)
        Dim i
        Dim data
        Dim codewordBitLength

        Dim segment

        For Each segment In m_segments.Items()
            Call bs.Append(segment.ModeIndicator, MODEINDICATOR_LENGTH)
            Call bs.Append(segment.CharCount, _
                           CharCountIndicator.GetLength( _
                                m_currVersion, segment.EncodingMode))

            data = segment.GetBytes()

            For i = 0 To UBound(data) - 1
                Call bs.Append(data(i), 8)
            Next

            codewordBitLength = segment.BitCount Mod 8

            If codewordBitLength = 0 Then
                codewordBitLength = 8
            End If

            Call bs.Append(data(UBound(data)) \ _
                           2 ^ (8 - codewordBitLength), codewordBitLength)
        Next
    End Sub

    Private Sub WriteTerminator(ByVal bs)
        Dim terminatorLength
        terminatorLength = m_dataBitCapacity - m_dataBitCounter

        If terminatorLength > MODEINDICATOR_LENGTH Then
            terminatorLength = MODEINDICATOR_LENGTH
        End If

        Call bs.Append(MODEINDICATOR_TERMINATOR_VALUE, terminatorLength)
    End Sub

    Private Sub WritePaddingBits(ByVal bs)
        If bs.Length Mod 8 > 0 Then
            Call bs.Append(&H0, 8 - (bs.Length Mod 8))
        End If
    End Sub

    Private Sub WritePadCodewords(ByVal bs)
        Dim numDataCodewords
        numDataCodewords = DataCodeword.GetTotalNumber( _
            m_parent.ErrorCorrectionLevel, m_currVersion)

        Dim flag
        flag = True

        Dim v

        Do While bs.Length < 8 * numDataCodewords
            If flag Then
                v = 236
            Else
                v = 17
            End If
            Call bs.Append(v, 8)
            flag = Not flag
        Loop
    End Sub

    Private Function GetModuleMatrix()
        Dim numModulesPerSide
        numModulesPerSide = Module.GetNumModulesPerSide(m_currVersion)

        Dim moduleMatrix
        ReDim moduleMatrix(numModulesPerSide - 1)

        Dim i
        Dim cols()

        For i = 0 To UBound(moduleMatrix)
            ReDim cols(numModulesPerSide - 1)
            moduleMatrix(i) = cols
        Next

        Call FinderPattern.Place(moduleMatrix)
        Call Separator.Place(moduleMatrix)
        Call TimingPattern.Place(moduleMatrix)

        If m_currVersion >= 2 Then
            Call AlignmentPattern.Place(m_currVersion, moduleMatrix)
        End If

        Call FormatInfo.PlaceTempBlank(moduleMatrix)

        If m_currVersion >= 7 Then
            Call VersionInfo.PlaceTempBlank(moduleMatrix)
        End If

        Call PlaceSymbolChar(moduleMatrix)
        Call RemainderBit.Place(moduleMatrix)

        Call Masking.Apply(m_currVersion, m_parent.ErrorCorrectionLevel, moduleMatrix)

        GetModuleMatrix = moduleMatrix
    End Function

    Private Sub PlaceSymbolChar(ByRef moduleMatrix())
        Dim data
        data = GetEncodingRegionBytes()

        Dim r
        r = UBound(moduleMatrix)

        Dim c
        c = UBound(moduleMatrix(0))

        Dim toLeft
        toLeft = True

        Dim rowDirection
        rowDirection = -1

        Dim bitPos
        Dim v

        For Each v In data
            bitPos = 7

            Do While bitPos >= 0
                If moduleMatrix(r)(c) = BLANK Then
                    If (v And 2 ^ bitPos) > 0 Then
                        moduleMatrix(r)(c) = WORD
                    Else
                        moduleMatrix(r)(c) = -WORD
                    End If

                    bitPos = bitPos - 1
                End If

                If toLeft Then
                    c = c - 1
                Else
                    If (r + rowDirection) < 0 Then
                        r = 0
                        rowDirection = 1
                        c = c - 1

                        If c = 6 Then
                            c = 5
                        End If

                    ElseIf ((r + rowDirection) > UBound(moduleMatrix)) Then
                        r = UBound(moduleMatrix)
                        rowDirection = -1
                        c = c - 1

                        If c = 6 Then
                            c = 5
                        End If

                    Else
                        r = r + rowDirection
                        c = c + 1
                    End If
                End If

                toLeft = Not toLeft
            Loop
        Next
    End Sub

    Private Function GetMonochromeBMP(ByVal moduleSize, ByVal foreColor, ByVal backColor)
        Dim foreRgb
        foreRgb = ColorCode.ToRGB(foreColor)
        Dim backRgb
        backRgb = ColorCode.ToRGB(backColor)

        If m_dataBitCounter = 0 Then Call Err.Raise(51)

        Dim moduleMatrix
        moduleMatrix = QuietZone.Place(GetModuleMatrix())

        Dim moduleCount
        moduleCount = UBound(moduleMatrix) + 1

        Dim pictWidth
        pictWidth = moduleCount * moduleSize

        Dim pictHeight
        pictHeight = moduleCount * moduleSize

        Dim rowBytesLen
        rowBytesLen = (pictWidth + 7) \ 8

        Dim pack8bit
        If pictWidth Mod 8 > 0 Then
            pack8bit = 8 - (pictWidth Mod 8)
        End If

        Dim pack32bit
        If rowBytesLen Mod 4 > 0 Then
            pack32bit = 8 * (4 - (rowBytesLen Mod 4))
        End If

        Dim rowSize
        rowSize = (pictWidth + pack8bit + pack32bit) \ 8

        Dim bitmapData
        Set bitmapData = New BinaryWriter

        Dim offset
        offset = 0

        Dim bs
        Set bs = New BitSequence

        Dim r, c
        Dim i
        Dim pixelColor
        Dim bitmapRow

        For r = UBound(moduleMatrix) To 0 Step -1
            Call bs.Clear

            For Each c In moduleMatrix(r)
                If IsDark(c) Then
                    pixelColor = 0
                Else
                    pixelColor = 1
                End If

                For i = 1 To moduleSize
                    Call bs.Append(pixelColor, 1)
                Next
            Next

            Call bs.Append(0, pack8bit)
            Call bs.Append(0, pack32bit)

            bitmapRow = bs.GetBytes()

            For i = 1 To moduleSize
                Call bitmapData.Append(bitmapRow)
            Next
        Next

        Dim ret
        Set ret = DIB.GetDIB(bitmapData, pictWidth, pictHeight, foreRgb, backRgb, True)

        Set GetMonochromeBMP = ret
    End Function

    Private Function GetTrueColorBMP(ByVal moduleSize, ByVal foreColor, ByVal backColor)
        Dim foreRgb
        foreRgb = ColorCode.ToRGB(foreColor)
        Dim backRgb
        backRgb = ColorCode.ToRGB(backColor)

        If m_dataBitCounter = 0 Then Call Err.Raise(51)

        Dim moduleMatrix
        moduleMatrix = QuietZone.Place(GetModuleMatrix())

        Dim pictWidth
        pictWidth = (UBound(moduleMatrix) + 1) * moduleSize

        Dim pictHeight
        pictHeight = pictWidth

        Dim rowBytesLen
        rowBytesLen = 3 * pictWidth

        Dim pack4byte
        If rowBytesLen Mod 4 > 0 Then
            pack4byte = 4 - (rowBytesLen Mod 4)
        End If

        Dim rowSize
        rowSize = rowBytesLen + pack4byte

        Dim bitmapData
        Set bitmapData = New BinaryWriter

        Dim offset
        offset = 0

        Dim r, c
        Dim i
        Dim colorRGB
        Dim bitmapRow()
        Dim idx

        For r = UBound(moduleMatrix) To 0 Step -1
            ReDim bitmapRow(rowSize - 1)
            idx = 0

            For Each c In moduleMatrix(r)
                If IsDark(c) Then
                    colorRGB = foreRgb
                Else
                    colorRGB = backRgb
                End If

                For i = 1 To moduleSize
                    bitmapRow(idx + 0) = CByte((colorRGB And &HFF0000) \ 2 ^ 16)
                    bitmapRow(idx + 1) = CByte((colorRGB And &HFF00&) \ 2 ^ 8)
                    bitmapRow(idx + 2) = CByte(colorRGB And &HFF&)
                    idx = idx + 3
                Next
            Next

            For i = 1 To pack4byte
                bitmapRow(idx) = CByte(0)
                idx = idx + 1
            Next

            For i = 1 To moduleSize
                Call bitmapData.Append(bitmapRow)
            Next
        Next

        Dim ret
        Set ret = DIB.GetDIB(bitmapData, pictWidth, pictHeight, foreRgb, backRgb, False)

        Set GetTrueColorBMP = ret
    End Function

    Private Function GetSvg(ByVal moduleSize, ByVal foreRgb)
        If m_dataBitCounter = 0 Then Call Err.Raise(51)

        If moduleSize < MIN_MODULE_SIZE Then Call Err.Raise(5)
        If ColorCode.IsWebColor(foreRgb) = False Then Call Err.Raise(5)

        Dim moduleMatrix
        moduleMatrix = QuietZone.Place(GetModuleMatrix())

        Dim imageWidth
        imageWidth = (UBound(moduleMatrix) + 1) * moduleSize

        Dim imageHeight
        imageHeight = imageWidth

        Dim img()
        ReDim img(imageHeight - 1)

        Dim imgRow()
        Dim r, c
        Dim i, j
        Dim v
        Dim cl

        r = 0
        Dim rowArray
        For Each rowArray In moduleMatrix
            ReDim imgRow(imageWidth - 1)
            c = 0
            For Each v In rowArray
                If IsDark(v) Then
                    cl = 1
                Else
                    cl = 0
                End If

                For j = 1 To moduleSize
                    imgRow(c) = cl
                    c = c + 1
                Next
            Next

            For i = 1 To moduleSize
                img(r) = imgRow
                r = r + 1
            Next
        Next

        Dim gpPaths
        gpPaths = GraphicPath.FindContours(img)

        Dim buf
        Set buf = New List

        Dim indent
        indent = String(5, " ")

        Dim gpPath
        Dim k
        For Each gpPath In gpPaths
            Call buf.Add(indent & "M ")

            For k = 0 To UBound(gpPath)
                Call buf.Add(CStr(gpPath(k).x) & "," & CStr(gpPath(k).y) & " ")
            Next
            Call buf.Add("Z" & vbNewLine)
        Next

        Dim data
        data = Trim(Join(buf.Items(), ""))
        data = Left(data, Len(data) - Len(vbNewLine))
        Dim svg
        svg = "<svg version=""1.1"" xmlns=""http://www.w3.org/2000/svg"" xmlns:xlink=""http://www.w3.org/1999/xlink""" & vbNewLine & _
              "  width=""" & CStr(imageWidth) & "px"" height=""" & CStr(imageHeight) & "px"" viewBox=""0 0 " & CStr(imageWidth) & " " & CStr(imageHeight) & """>" & vbNewLine & _
              "<path fill=""" & foreRgb & """ stroke=""" & foreRgb & """ stroke-width=""1""" & vbNewLine & _
              "  d=""" & data & """ />" & vbNewLine & _
              "</svg>"

        GetSvg = svg
    End Function

    Private Function GetMonochromePNG(ByVal moduleSize, ByVal foreRgb, ByVal backRgb)
        Dim foreColorRgb
        foreColorRgb = ColorCode.ToRGB(foreRgb)
        Dim backColorRgb
        backColorRgb = ColorCode.ToRGB(backRgb)

        If m_dataBitCounter = 0 Then Call Err.Raise(51)

        Dim moduleMatrix
        moduleMatrix = QuietZone.Place(GetModuleMatrix())

        Dim moduleCount
        moduleCount = UBound(moduleMatrix) + 1

        Dim pictWidth
        pictWidth = moduleCount * moduleSize

        Dim pictHeight
        pictHeight = moduleCount * moduleSize

        Dim pack8bit
        If pictWidth Mod 8 > 0 Then
            pack8bit = 8 - (pictWidth Mod 8)
        End If

        Dim bs
        Set bs = New BitSequence

        Dim pixelColor

        Dim filterType
        filterType = 0

        Dim r, c
        Dim i, j
        For r = 0 To UBound(moduleMatrix)
            For i = 1 To moduleSize
                Call bs.Append(filterType, 8)

                For Each c In moduleMatrix(r)
                    If IsDark(c) Then
                        pixelColor = 0
                    Else
                        pixelColor = 1
                    End If

                    For j = 1 To moduleSize
                        Call bs.Append(pixelColor, 1)
                    Next
                Next
                Call bs.Append(0, pack8bit)
            Next
        Next

        Dim bitmapData
        bitmapData = bs.GetBytes()

        Set GetMonochromePNG = PNG.GetPNG( _
            bitmapData, pictWidth, pictHeight, foreColorRgb, backColorRgb, _
            PngColorType_pIndexColor _
        )
    End Function

    Private Function GetTrueColorPNG( _
        ByVal moduleSize, ByVal foreRgb, ByVal backRgb, ByVal bkStyle)

        Dim foreColorRgb
        foreColorRgb = ColorCode.ToRGB(foreRgb)
        Dim backColorRgb
        backColorRgb = ColorCode.ToRGB(backRgb)

        If m_dataBitCounter = 0 Then Call Err.Raise(51)

        Dim moduleMatrix
        moduleMatrix = QuietZone.Place(GetModuleMatrix())

        Dim pictWidth
        pictWidth = (UBound(moduleMatrix) + 1) * moduleSize

        Dim pictHeight
        pictHeight = pictWidth

        Dim rowSize
        If bkStyle = BackStyle_bkTransparent Then
            rowSize = 1 + 4 * pictWidth
        Else
            rowSize = 1 + 3 * pictWidth
        End If

        Dim bitmapData
        ReDim bitmapData(rowSize * pictHeight - 1)

        Dim offset
        offset = 0

        Dim colorRgb
        Dim alpha
        Dim idx
        idx = 0

        Dim filterType
        filterType = 0

        Dim r, c
        Dim i, j
        For r = 0 To UBound(moduleMatrix)
            For i = 1 To moduleSize
                bitmapData(idx) = CByte(filterType)
                idx = idx + 1

                For Each c In moduleMatrix(r)
                    If IsDark(c) Then
                        colorRgb = foreColorRgb
                        alpha = CByte(&HFF)
                    Else
                        colorRgb = backColorRgb
                        alpha = CByte(0)
                    End If

                    For j = 1 To moduleSize
                        bitmapData(idx + 0) = CByte(colorRgb And &HFF&)               ' R
                        bitmapData(idx + 1) = CByte((colorRgb And &HFF00&) \ 2 ^ 8)   ' G
                        bitmapData(idx + 2) = CByte((colorRgb And &HFF0000) \ 2 ^ 16) ' B
                        idx = idx + 3

                        If bkStyle = BackStyle_bkTransparent Then
                            bitmapData(idx) = alpha
                            idx = idx + 1
                        End If
                    Next
                Next
            Next
        Next

        Dim tColor
        If bkStyle = BackStyle_bkTransparent Then
            tColor = PngColorType_pTrueColorAlpha
        Else
            tColor = PngColorType_pTrueColor
        End If

        Set GetTrueColorPNG = PNG.GetPNG( _
            bitmapData, pictWidth, pictHeight, foreColorRgb, backColorRgb, tColor)
    End Function

    Public Sub SaveAs(ByVal filename)
        Call SaveAs2(filename, 5, False, False, "#000000", "#FFFFFF")
    End Sub

    Public Sub SaveAs2( _
        ByVal filename, ByVal moduleSize, ByVal monochrome, ByVal transparent, _
        ByVal foreRgb, ByVal backRgb)

        If m_dataBitCounter = 0 Then Call Err.Raise(51)

        If Len(filename) = 0 Then Call Err.Raise(5, , "[filename]")
        If moduleSize < MIN_MODULE_SIZE Then Call Err.Raise(5, , "[moduleSize]")
        If VarType(monochrome) <> vbBoolean then  Call Err.Raise(5, , "[monochrome]")
        If ColorCode.IsWebColor(foreRgb) = False Then Call Err.Raise(5, , "[foreRgb]")
        If ColorCode.IsWebColor(backRgb) = False Then Call Err.Raise(5, , "[backRgb]")

        Dim fso
        Set fso = CreateObject("Scripting.FileSystemObject")

        Dim ext
        ext = fso.GetExtensionName(filename)

        Dim bw, txt
        Dim ts

        Select Case LCase(ext)
            Case "bmp"
                If monochrome Then
                    Set bw = GetMonochromeBMP(moduleSize, foreRgb, backRgb)
                Else
                    Set bw = GetTrueColorBMP(moduleSize, foreRgb, backRgb)
                End If
                Call bw.SaveToFile(filename, adSaveCreateOverWrite)
            Case "png"
                If monochrome Then
                    Set bw = GetMonochromePNG(moduleSize, foreRgb, backRgb)
                Else
                    If transparent Then
                        Set bw = GetTrueColorPNG(moduleSize, foreRgb, backRgb, BackStyle_bkTransparent)
                    Else
                        Set bw = GetTrueColorPNG(moduleSize, foreRgb, backRgb, BackStyle_bkOpaque)
                    End If
                End If
                Call bw.SaveToFile(filename, adSaveCreateOverWrite)
            Case "svg"
                txt = GetSvg(moduleSize, foreRgb)
                Set ts = fso.CreateTextFile(filename, True)
                Call ts.WriteLine(txt)
                Call ts.Close
            Case Else
                Call Err.Raise(5)
        End Select
    End Sub
End Class


Class Symbols
    Private m_items

    Private m_minVersion
    Private m_maxVersion
    Private m_errorCorrectionLevel
    Private m_structuredAppendAllowed
    Private m_byteModeCharsetName

    Private m_parity

    Private m_currSymbol

    Private m_encNum
    Private m_encAlpha
    Private m_encByte
    Private m_encKanji

    Public Sub Init(ByVal ecLevel, ByVal maxVer, ByVal allowStructuredAppend)
        If Not (MIN_VERSION <= maxVer And maxVer <= MAX_VERSION) Then
            Call Err.Raise(5)
        End If

        Set m_items = New List

        Set m_encNum = CreateEncoder(MODE_NUMERIC)
        Set m_encAlpha = CreateEncoder(MODE_ALPHA_NUMERIC)
        Set m_encByte = CreateEncoder(MODE_BYTE)
        Set m_encKanji = CreateEncoder(MODE_KANJI)

        m_minVersion = 1
        m_maxVersion = maxVer
        m_errorCorrectionLevel = ecLevel
        m_structuredAppendAllowed = allowStructuredAppend

        m_parity = 0

        Set m_currSymbol = New Symbol
        Call m_currSymbol.Init(Me)
        Call m_items.Add(m_currSymbol)
    End Sub

    Public Property Get Item(ByVal idx)
        Set Item = m_items.Item(idx)
    End Property

    Public Property Get Count()
        Count = m_items.Count
    End Property

    Public Property Get StructuredAppendAllowed()
        StructuredAppendAllowed = m_structuredAppendAllowed
    End Property

    Public Property Get Parity()
        Parity = m_parity
    End Property

    Public Property Get MinVersion()
        MinVersion = m_minVersion
    End Property
    Public Property Let MinVersion(ByVal Value)
        m_minVersion = Value
    End Property

    Public Property Get MaxVersion()
        MaxVersion = m_maxVersion
    End Property

    Public Property Get ErrorCorrectionLevel()
        ErrorCorrectionLevel = m_errorCorrectionLevel
    End Property

    Private Function Add()
        Set m_currSymbol = New Symbol
        Call m_currSymbol.Init(Me)
        Call m_items.Add(m_currSymbol)

        Set Add = m_currSymbol
    End Function

    Public Sub AppendText(ByVal s)
        If Len(s) = 0 Then Call Err.Raise(5)

        Dim oldMode
        Dim newMode
        Dim i
        For i = 1 To Len(s)
            oldMode = m_currSymbol.CurrentEncodingMode

            Select Case oldMode
                Case MODE_UNKNOWN
                    newMode = SelectInitialMode(s, i)
                Case MODE_NUMERIC
                    newMode = SelectModeWhileInNumeric(s, i)
                Case MODE_ALPHA_NUMERIC
                    newMode = SelectModeWhileInAlphanumeric(s, i)
                Case MODE_BYTE
                    newMode = SelectModeWhileInByte(s, i)
                Case MODE_KANJI
                    newMode = SelectInitialMode(s, i)
                Case Else
                    Call Err.Raise(51)
            End Select

            If newMode <> oldMode Then
                If Not m_currSymbol.TrySetEncodingMode(newMode, Mid(s, i, 1)) Then
                    If Not m_structuredAppendAllowed Or m_items.Count = 16 Then
                        Call Err.Raise(6)
                    End If

                    Call Add
                    newMode = SelectInitialMode(s, i)
                    Call m_currSymbol.TrySetEncodingMode(newMode, Mid(s, i, 1))
                End If
            End If

            If Not m_currSymbol.TryAppend(Mid(s, i, 1)) Then
                If Not m_structuredAppendAllowed Or m_items.Count = 16 Then
                    Call Err.Raise(6)
                End If

                Call Add
                newMode = SelectInitialMode(s, i)
                Call m_currSymbol.TrySetEncodingMode(newMode, Mid(s, i, 1))
                Call m_currSymbol.TryAppend(Mid(s, i, 1))
            End If
        Next
    End Sub

    Public Sub UpdateParity(ByVal c)
        Dim code
        code = Asc(c) And &HFFFF&

        Dim msb
        Dim lsb

        msb = (code And &HFF00&) \ 2 ^ 8
        lsb = code And &HFF&

        If msb > 0 Then
            m_parity = m_parity Xor msb
        End If

        m_parity = m_parity Xor lsb
    End Sub

    Private Function SelectInitialMode(ByRef s, ByVal startIndex)
        If m_encKanji.InSubset(Mid(s, startIndex, 1)) Then
            SelectInitialMode = MODE_KANJI
            Exit Function
        End If

        If m_encByte.InExclusiveSubset(Mid(s, startIndex, 1)) Then
            SelectInitialMode = MODE_BYTE
            Exit Function
        End If

        If m_encAlpha.InExclusiveSubset(Mid(s, startIndex, 1)) Then
            SelectInitialMode = SelectModeWhenInitialDataAlphaNumeric(s, startIndex)
            Exit Function
        End If

        If m_encNum.InSubset(Mid(s, startIndex, 1)) Then
            SelectInitialMode = SelectModeWhenInitialDataNumeric(s, startIndex)
            Exit Function
        End If

        Call Err.Raise(51)
    End Function

    Private Function SelectModeWhenInitialDataAlphaNumeric(ByRef s, ByVal startIndex)
        Dim cnt
        cnt = 0

        Dim i
        For i = startIndex To Len(s)
            If m_encAlpha.InExclusiveSubset(Mid(s, i, 1)) Then
                cnt = cnt + 1
            Else
                Exit For
            End If
        Next

        Dim flg
        flg = False

        Dim ver
        ver = m_currSymbol.Version

        If 1 <= ver And ver <= 9 Then
            flg = cnt < 6
        ElseIf 10 <= ver And ver <= 26 Then
            flg = cnt < 7
        ElseIf 27 <= ver And ver <= 40 Then
            flg = cnt < 8
        Else
            Call Err.Raise(51)
        End If

        If flg Then
            If (startIndex + cnt) <= Len(s) Then
                If m_encByte.InSubset(Mid(s, startIndex + cnt, 1)) Then
                    SelectModeWhenInitialDataAlphaNumeric = MODE_BYTE
                    Exit Function
                End If
            End If
        End If

        SelectModeWhenInitialDataAlphaNumeric = MODE_ALPHA_NUMERIC
    End Function

    Private Function SelectModeWhenInitialDataNumeric(ByRef s, ByVal startIndex)
        Dim cnt
        cnt = 0

        Dim i
        For i = startIndex To Len(s)
            If m_encNum.InSubset(Mid(s, i, 1)) Then
                cnt = cnt + 1
            Else
                Exit For
            End If
        Next

        Dim flg

        Dim ver
        ver = m_currSymbol.Version

        If 1 <= ver And ver <= 9 Then
            flg = cnt < 4
        ElseIf 10 <= ver And ver <= 26 Then
            flg = cnt < 4
        ElseIf 27 <= ver And ver <= 40 Then
            flg = cnt < 5
        Else
            Call Err.Raise(51)
        End If

        If flg Then
            If (startIndex + cnt) <= Len(s) Then
                SelectModeWhenInitialDataNumeric = MODE_BYTE
                Exit Function
            End If
        End If

        If 1 <= ver And ver <= 9 Then
            flg = cnt < 7
        ElseIf 10 <= ver And ver <= 26 Then
            flg = cnt < 8
        ElseIf 27 <= ver And ver <= 40 Then
            flg = cnt < 9
        Else
            Call Err.Raise(51)
        End If

        If flg Then
            If (startIndex + cnt) <= Len(s) Then
                SelectModeWhenInitialDataNumeric = MODE_ALPHA_NUMERIC
                Exit Function
            End If
        End If

        SelectModeWhenInitialDataNumeric = MODE_NUMERIC
    End Function

    Private Function SelectModeWhileInNumeric(ByRef s, ByVal startIndex)
        If m_encKanji.InSubset(Mid(s, startIndex, 1)) Then
            SelectModeWhileInNumeric = MODE_KANJI
            Exit Function
        End If

        If m_encByte.InExclusiveSubset(Mid(s, startIndex, 1)) Then
            SelectModeWhileInNumeric = MODE_BYTE
            Exit Function
        End If

        If m_encAlpha.InExclusiveSubset(Mid(s, startIndex, 1)) Then
            SelectModeWhileInNumeric = MODE_ALPHA_NUMERIC
            Exit Function
        End If

        SelectModeWhileInNumeric = MODE_NUMERIC
    End Function

    Private Function SelectModeWhileInAlphanumeric(ByRef s, ByVal startIndex)
        If m_encKanji.InSubset(Mid(s, startIndex, 1)) Then
            SelectModeWhileInAlphanumeric = MODE_KANJI
            Exit Function
        End If

        If m_encByte.InExclusiveSubset(Mid(s, startIndex, 1)) Then
            SelectModeWhileInAlphanumeric = MODE_BYTE
            Exit Function
        End If

        If MustChangeAlphanumericToNumeric(s, startIndex) Then
            SelectModeWhileInAlphanumeric = MODE_NUMERIC
            Exit Function
        End If

        SelectModeWhileInAlphanumeric = MODE_ALPHA_NUMERIC
    End Function

    Private Function MustChangeAlphanumericToNumeric(ByRef s, ByVal startIndex)
        Dim cnt
        cnt = 0

        Dim ret
        ret = False

        Dim i
        For i = startIndex To Len(s)
            If Not m_encAlpha.InSubset(Mid(s, i, 1)) Then
                Exit For
            End If

            If m_encNum.InSubset(Mid(s, i, 1)) Then
                cnt = cnt + 1
            Else
                ret = True
                Exit For
            End If
        Next

        Dim ver
        ver = m_currSymbol.Version

        If ret Then
            If 1 <= ver And ver <= 9 Then
                ret = cnt >= 13
            ElseIf 10 <= ver And ver <= 26 Then
                ret = cnt >= 15
            ElseIf 27 <= ver And ver <= 40 Then
                ret = cnt >= 17
            Else
                Call Err.Raise(51)
            End If
        End If

        MustChangeAlphanumericToNumeric = ret
    End Function

    Private Function SelectModeWhileInByte(ByRef s, ByVal startIndex)
        If m_encKanji.InSubset(Mid(s, startIndex, 1)) Then
            SelectModeWhileInByte = MODE_KANJI
            Exit Function
        End If

        If MustChangeByteToNumeric(s, startIndex) Then
            SelectModeWhileInByte = MODE_NUMERIC
            Exit Function
        End If

        If MustChangeByteToAlphanumeric(s, startIndex) Then
            SelectModeWhileInByte = MODE_ALPHA_NUMERIC
            Exit Function
        End If

        SelectModeWhileInByte = MODE_BYTE
    End Function

    Private Function MustChangeByteToNumeric(ByRef s, ByVal startIndex)
        Dim cnt
        cnt = 0

        Dim ret
        ret = False

        Dim i
        For i = startIndex To Len(s)
            If Not m_encByte.InSubset(Mid(s, i, 1)) Then
                Exit For
            End If

            If m_encNum.InSubset(Mid(s, i, 1)) Then
                cnt = cnt + 1
            ElseIf m_encByte.InExclusiveSubset(Mid(s, i, 1)) Then
                ret = True
                Exit For
            Else
                Exit For
            End If
        Next

        Dim ver
        ver = m_currSymbol.Version

        If ret Then
            If 1 <= ver And ver <= 9 Then
                ret = cnt >= 6
            ElseIf 10 <= ver And ver <= 26 Then
                ret = cnt >= 8
            ElseIf 27 <= ver And ver <= 40 Then
                ret = cnt >= 9
            Else
                Call Err.Raise(51)
            End If
        End If

        MustChangeByteToNumeric = ret
    End Function

    Private Function MustChangeByteToAlphanumeric(ByRef s, ByVal startIndex)
        Dim ret

        Dim cnt
        cnt = 0

        Dim i
        For i = startIndex To Len(s)
            If Not m_encByte.InSubset(Mid(s, i, 1)) Then
                Exit For
            End If

            If m_encAlpha.InExclusiveSubset(Mid(s, i, 1)) Then
                cnt = cnt + 1
            ElseIf m_encByte.InExclusiveSubset(Mid(s, i, 1)) Then
                ret = True
                Exit For
            Else
                Exit For
            End If
        Next

        Dim ver
        ver = m_currSymbol.Version

        If ret Then
            If 1 <= ver And ver <= 9 Then
                ret = cnt >= 11
            ElseIf 10 <= ver And ver <= 26 Then
                ret = cnt >= 15
            ElseIf 27 <= ver And ver <= 40 Then
                ret = cnt >= 16
            Else
                Call Err.Raise(51)
            End If
        End If

        MustChangeByteToAlphanumeric = ret
    End Function
End Class


Class TimingPattern_
    Public Sub Place(ByRef moduleMatrix())
        Dim i
        Dim v

        For i = 8 To UBound(moduleMatrix) - 8
            If i Mod 2 = 0 Then
                v = TIMING_PTN
            Else
                v = -TIMING_PTN
            End If

            moduleMatrix(6)(i) = v
            moduleMatrix(i)(6) = v
        Next
    End Sub
End Class


Class VersionInfo_
    Private m_versionInfoValues

    Private Sub Class_Initialize()
        m_versionInfoValues = Array( _
            -1, -1, -1, -1, -1, -1, -1, _
            &H7C94&, &H85BC&, &H9A99&, &HA4D3&, &HBBF6&, &HC762&, &HD847&, &HE60D&, _
            &HF928&, &H10B78, &H1145D, &H12A17, &H13532, &H149A6, &H15683, &H168C9, _
            &H177EC, &H18EC4, &H191E1, &H1AFAB, &H1B08E, &H1CC1A, &H1D33F, &H1ED75, _
            &H1F250, &H209D5, &H216F0, &H228BA, &H2379F, &H24B0B, &H2542E, &H26A64, _
            &H27541, &H28C69 _
        )
    End Sub

    Public Sub Place(ByVal ver, ByRef moduleMatrix())
        Dim numModulesPerSide
        numModulesPerSide = UBound(moduleMatrix) + 1

        Dim versionInfoValue
        versionInfoValue = m_versionInfoValues(ver)

        Dim p1
        p1 = 0

        Dim p2
        p2 = numModulesPerSide - 11

        Dim i
        Dim v

        For i = 0 To 17
            If (versionInfoValue And 2 ^ i) > 0 Then
                v = VERSION_INFO
            Else
                v = -VERSION_INFO
            End IF

            moduleMatrix(p1)(p2) = v
            moduleMatrix(p2)(p1) = v

            p2 = p2 + 1

            If i Mod 3 = 2 Then
                p1 = p1 + 1
                p2 = numModulesPerSide - 11
            End If
        Next
    End Sub

    Public Sub PlaceTempBlank(ByRef moduleMatrix())
        Dim numModulesPerSide
        numModulesPerSide = UBound(moduleMatrix) + 1

        Dim i, j

        For i = 0 To 5
            For j = numModulesPerSide - 11 To numModulesPerSide - 9
                moduleMatrix(i)(j) = -VERSION_INFO
                moduleMatrix(j)(i) = -VERSION_INFO
            Next
        Next
    End Sub
End Class


Class BitmapFileHeader
    Private m_bfType
    Public Property Let bfType(ByVal Value)
        m_bfType = CInt(Value)
    End Property
    Public Property Get bfType()
        bfType = m_bfType
    End Property

    Private m_bfSize
    Public Property Let bfSize(ByVal Value)
        m_bfSize = CLng(Value)
    End Property
    Public Property Get bfSize()
        bfSize = m_bfSize
    End Property

    Private m_bfReserved1
    Public Property Let bfReserved1(ByVal Value)
        m_bfReserved1 = CInt(Value)
    End Property
    Public Property Get bfReserved1()
        bfReserved1 = m_bfReserved1
    End Property

    Private m_bfReserved2
    Public Property Let bfReserved2(ByVal Value)
        m_bfReserved2 = CInt(Value)
    End Property
    Public Property Get bfReserved2()
        bfReserved2 = m_bfReserved2
    End Property

    Private m_bfOffBits
    Public Property Let bfOffBits(ByVal Value)
        m_bfOffBits = CLng(Value)
    End Property
    Public Property Get bfOffBits()
        bfOffBits = m_bfOffBits
    End Property
End Class


Class BitmapInfoHeader
    Private m_biSize
    Public Property Let biSize(ByVal Value)
        m_biSize = CLng(Value)
    End Property
    Public Property Get biSize()
        biSize = m_biSize
    End Property

    Private m_biWidth
    Public Property Let biWidth(ByVal Value)
        m_biWidth = CLng(Value)
    End Property
    Public Property Get biWidth()
        biWidth = m_biWidth
    End Property

    Private m_biHeight
    Public Property Let biHeight(ByVal Value)
        m_biHeight = CLng(Value)
    End Property
    Public Property Get biHeight()
        biHeight = m_biHeight
    End Property

    Private m_biPlanes
    Public Property Let biPlanes(ByVal Value)
        m_biPlanes = CInt(Value)
    End Property
    Public Property Get biPlanes()
        biPlanes = m_biPlanes
    End Property

    Private m_biBitCount
    Public Property Let biBitCount(ByVal Value)
        m_biBitCount = CInt(Value)
    End Property
    Public Property Get biBitCount()
        biBitCount = m_biBitCount
    End Property

    Private m_biCompression
    Public Property Let biCompression(ByVal Value)
        m_biCompression = CLng(Value)
    End Property
    Public Property Get biCompression()
        biCompression = m_biCompression
    End Property

    Private m_biSizeImage
    Public Property Let biSizeImage(ByVal Value)
        m_biSizeImage = CLng(Value)
    End Property
    Public Property Get biSizeImage()
        biSizeImage = m_biSizeImage
    End Property

    Private m_biXPelsPerMeter
    Public Property Let biXPelsPerMeter(ByVal Value)
        m_biXPelsPerMeter = CLng(Value)
    End Property
    Public Property Get biXPelsPerMeter()
        biXPelsPerMeter = m_biXPelsPerMeter
    End Property

    Private m_biYPelsPerMeter
    Public Property Let biYPelsPerMeter(ByVal Value)
        m_biYPelsPerMeter = CLng(Value)
    End Property
    Public Property Get biYPelsPerMeter()
        biYPelsPerMeter = m_biYPelsPerMeter
    End Property

    Private m_biClrUsed
    Public Property Let biClrUsed(ByVal Value)
        m_biClrUsed = CLng(Value)
    End Property
    Public Property Get biClrUsed()
        biClrUsed = m_biClrUsed
    End Property

    Private m_biClrImportant
    Public Property Let biClrImportant(ByVal Value)
        m_biClrImportant = CLng(Value)
    End Property
    Public Property Get biClrImportant()
        biClrImportant = m_biClrImportant
    End Property
End Class


Class RgbQuad
    Private m_rgbBlue
    Public Property Let rgbBlue(ByVal Value)
        m_rgbBlue = CByte(Value)
    End Property
    Public Property Get rgbBlue()
        rgbBlue = m_rgbBlue
    End Property

    Private m_rgbGreen
    Public Property Let rgbGreen(ByVal Value)
        m_rgbGreen = CByte(Value)
    End Property
    Public Property Get rgbGreen()
        rgbGreen = m_rgbGreen
    End Property

    Private m_rgbRed
    Public Property Let rgbRed(ByVal Value)
        m_rgbRed = CByte(Value)
    End Property
    Public Property Get rgbRed()
        rgbRed = m_rgbRed
    End Property

    Private m_rgbReserved
    Public Property Let rgbReserved(ByVal Value)
        m_rgbReserved = CByte(Value)
    End Property
    Public Property Get rgbReserved()
        rgbReserved = m_rgbReserved
    End Property
End Class


Class ColorCode_
    Public Property Get BLACK()
        BLACK = "#000000"
    End Property

    Public Property Get WHITE()
        WHITE = "#FFFFFF"
    End Property

    Public Function IsWebColor(arg)
        Dim re
        Set re = CreateObject("VBScript.RegExp")
        re.Pattern = "^#[0-9A-Fa-f]{6}$"
        Dim ret
        ret = re.Test(arg)
        IsWebColor = ret
    End Function

    Public Function ToRGB(ByVal arg)
        If Not IsWebColor(arg) Then Call Err.Raise(5)

        Dim ret
        ret = RGB(CInt("&h" & Mid(arg, 2, 2)), _
                  CInt("&h" & Mid(arg, 4, 2)), _
                  CInt("&h" & Mid(arg, 6, 2)))

        ToRGB = ret
    End Function
End Class


Class Point
    Private m_x
    Private m_y

    Public Sub Init(ByVal x, ByVal y)
        m_x = x
        m_y = y
    End Sub

    Public Property Get x()
        x = m_x
    End Property
    Public Property Let x(ByVal Value)
        m_x = Value
    End Property

    Public Property Get y()
        y = m_y
    End Property
    Public Property Let y(ByVal Value)
        m_y = Value
    End Property

    Public Function Clone()
        Dim ret
        Set ret = New Point
        Call ret.Init(x, y)

        Set Clone = ret
    End Function

    Public Function Equals(ByVal obj)
        Equals = (x = obj.x) And (y = obj.y)
    End Function
End Class


Class DIB_
    Public Function GetDIB( _
      ByVal bitmapData, ByVal pictWidth, ByVal pictHeight, ByVal foreRgb, ByVal backRgb, ByVal monochrome)
        Const BF_SIZE = 14
        Const BI_SIZE = 40

        Dim bfOffBits
        Dim biBitCount

        Dim palette()

        If Not monochrome Then
            biBitCount = 24
            bfOffBits = BF_SIZE + BI_SIZE
        Else
            ReDim palette(1)
            Set palette(0) = New RgbQuad
            Set palette(1) = New RgbQuad

            With palette(0)
                .rgbBlue = (foreRgb And &HFF0000) \ 2 ^ 16
                .rgbGreen = (foreRgb And &HFF00&) \ 2 ^ 8
                .rgbRed = foreRgb And &HFF&
                .rgbReserved = 0
            End With

            With palette(1)
                .rgbBlue = (backRgb And &HFF0000) \ 2 ^ 16
                .rgbGreen = (backRgb And &HFF00&) \ 2 ^ 8
                .rgbRed = backRgb And &HFF&
                .rgbReserved = 0
            End With

            biBitCount = 1
            bfOffBits = BF_SIZE + BI_SIZE + (4 * (UBound(palette) + 1))
        End If

        Dim bfh
        Set bfh = New BitmapFileHeader
        With bfh
            .bfType = &H4D42
            .bfSize = bfOffBits + bitmapData.Size
            .bfReserved1 = 0
            .bfReserved2 = 0
            .bfOffBits = bfOffBits
        End With

        Dim bih
        Set bih = New BitmapInfoHeader
        With bih
            .biSize = BI_SIZE
            .biWidth = pictWidth
            .biHeight = pictHeight
            .biPlanes = 1
            .biBitCount = biBitCount
            .biCompression = 0
            .biSizeImage = 0
            .biXPelsPerMeter = 0
            .biYPelsPerMeter = 0
            .biClrUsed = 0
            .biClrImportant = 0
        End With

        Dim ret
        Set ret = New BinaryWriter

        With bfh
            Call ret.Append(.bfType)
            Call ret.Append(.bfSize)
            Call ret.Append(.bfReserved1)
            Call ret.Append(.bfReserved2)
            Call ret.Append(.bfOffBits)
        End With

        With bih
            Call ret.Append(.biSize)
            Call ret.Append(.biWidth)
            Call ret.Append(.biHeight)
            Call ret.Append(.biPlanes)
            Call ret.Append(.biBitCount)
            Call ret.Append(.biCompression)
            Call ret.Append(.biSizeImage)
            Call ret.Append(.biXPelsPerMeter)
            Call ret.Append(.biYPelsPerMeter)
            Call ret.Append(.biClrUsed)
            Call ret.Append(.biClrImportant)
        End With

        Dim i
        If monochrome Then
            For i = 0 To UBound(palette)
                With palette(i)
                    Call ret.Append(.rgbBlue)
                    Call ret.Append(.rgbGreen)
                    Call ret.Append(.rgbRed)
                    Call ret.Append(.rgbReserved)
                End With
            Next
        End If

        Call bitmapData.CopyTo(ret)

        Set GetDIB = ret
    End Function
End Class


Class GraphicPath_
    Public Function FindContours(ByRef img)
        Dim MAX_VALUE
        MAX_VALUE = &H7FFFFFFF

        Dim gpPaths
        Set gpPaths = New List

        Dim st, dr
        Dim x, y
        Dim p
        Dim gpPath

        For y = 0 To UBound(img) - 1
            For x = 0 To UBound(img(y)) - 1
                If Not (img(y)(x) = MAX_VALUE) And _
                    (img(y)(x) > 0 And img(y)(x + 1) <= 0) Then

                    img(y)(x) = MAX_VALUE
                    Set st = New Point
                    Call st.Init(x, y)
                    Set gpPath = New List
                    Call gpPath.Add(st)

                    dr = DIRECTION_UP
                    Set p = st.Clone()
                    p.y = p.y - 1

                    Do Until p.Equals(st)
                        Select Case dr
                            Case DIRECTION_UP
                                If img(p.y)(p.x) > 0 Then
                                    img(p.y)(p.x) = MAX_VALUE

                                    If img(p.y)(p.x + 1) <= 0 Then
                                        Set p = p.Clone()
                                        p.y = p.y - 1
                                    Else
                                        Call gpPath.Add(p)
                                        dr = DIRECTION_RIGHT
                                        Set p = p.Clone()
                                        p.x = p.x + 1
                                    End If
                                Else
                                    Set p = p.Clone()
                                    p.y = p.y + 1
                                    Call gpPath.Add(p)

                                    dr = DIRECTION_LEFT
                                    Set p = p.Clone()
                                    p.x = p.x - 1
                                End If

                            Case DIRECTION_DOWN
                                If img(p.y)(p.x) > 0 Then
                                    img(p.y)(p.x) = MAX_VALUE

                                    If img(p.y)(p.x - 1) <= 0 Then
                                        Set p = p.Clone()
                                        p.y = p.y + 1
                                    Else
                                        Call gpPath.Add(p)

                                        dr = DIRECTION_LEFT
                                        Set p = p.Clone()
                                        p.x = p.x - 1
                                    End If
                                Else
                                    Set p = p.Clone()
                                    p.y = p.y - 1
                                    Call gpPath.Add(p)

                                    dr = DIRECTION_RIGHT
                                    Set p = p.Clone()
                                    p.x = p.x + 1
                                End If

                            Case DIRECTION_LEFT
                                If img(p.y)(p.x) > 0 Then
                                    img(p.y)(p.x) = MAX_VALUE

                                    If img(p.y - 1)(p.x) <= 0 Then
                                        Set p = p.Clone()
                                        p.x = p.x - 1
                                    Else
                                        Call gpPath.Add(p)

                                        dr = DIRECTION_UP
                                        Set p = p.Clone()
                                        p.y = p.y - 1
                                    End If
                                Else
                                    Set p = p.Clone()
                                    p.x = p.x + 1
                                    Call gpPath.Add(p)

                                    dr = DIRECTION_DOWN
                                    Set p = p.Clone()
                                    p.y = p.y + 1
                                End If

                            Case DIRECTION_RIGHT
                                If img(p.y)(p.x) > 0 Then
                                    img(p.y)(p.x) = MAX_VALUE

                                    If img(p.y + 1)(p.x) <= 0 Then
                                        Set p = p.Clone()
                                        p.x = p.x + 1
                                    Else
                                        Call gpPath.Add(p)

                                        dr = DIRECTION_DOWN
                                        Set p = p.Clone()
                                        p.y = p.y + 1
                                    End If
                                Else
                                    Set p = p.Clone()
                                    p.x = p.x - 1
                                    Call gpPath.Add(p)

                                    dr = DIRECTION_UP
                                    Set p = p.Clone()
                                    p.y = p.y - 1
                                End If
                            Case Else
                                Call Err.Raise(51)
                        End Select
                    Loop

                    Call gpPaths.Add(gpPath.Items())
                End If
            Next
        Next

        FindContours = gpPaths.Items()
    End Function
End Class


Class CRC32_
    Private m_crcTable(255)
    Private m_crcTableComputed

    Public Function Checksum(ByRef data)
        Checksum = Update(0, data)
    End Function

    Public Function Update(ByVal crc, ByRef data)
        Dim c
        c = crc Xor &HFFFFFFFF

        If Not m_crcTableComputed Then
            Call MakeCrcTable
        End If

        Dim n
        For n = 0 To UBound(data)
            c = m_crcTable((c Xor data(n)) And &HFF) Xor _
                    (((c And &HFFFFFF00) \ 2 ^ 8) And &HFFFFFF)
        Next

        Update = c Xor &HFFFFFFFF
    End Function

    Private Sub MakeCrcTable()
        Dim c

        Dim k
        Dim n
        For n = 0 To 255
            c = n
            For k = 0 To 7
                If c And 1 Then
                    c = &HEDB88320 Xor ((c And &HFFFFFFFE) \ 2 And &H7FFFFFFF)
                Else
                    c = (c \ 2) And &H7FFFFFFF
                End If
            Next

            m_crcTable(n) = c
        Next

        m_crcTableComputed = True
    End Sub
End Class


Class ADLER32_
    Public Function Checksum(ByRef data)
        Checksum = Update(1, data)
    End Function

    Public Function Update(ByVal adler, ByRef data)
        Dim s1
        s1 = adler And &HFFFF&

        Dim s2
        s2 = (adler \ 2 ^ 16) And &HFFFF&

        Dim n
        For n = 0 To UBound(data)
            s1 = (s1 + data(n)) Mod 65521
            s2 = (s2 + s1) Mod 65521
        Next

        Dim temp

        If (s2 And &H8000&) > 0 Then
            temp = ((s2 And &H7FFF&) * 2 ^ 16) Or &H80000000
        Else
            temp = s2 * 2 ^ 16
        End If

        Update = temp + s1
    End Function
End Class


Class Deflate_
    Public Function Compress(ByRef data, ByVal btype)
        If btype <> DeflateBType_NoCompression Then Call Err.Raise(5)

        Dim bytesLen
        bytesLen = UBound(data) + 1

        Dim quotient
        quotient = bytesLen \ &HFFFF&

        Dim remainder
        remainder = bytesLen Mod &HFFFF&

        Dim bufferSize
        bufferSize = quotient * (1 + 4 + &HFFFF&)

        If remainder > 0 Then
            bufferSize = bufferSize + (1 + 4 + remainder)
        End If

        ReDim ret(bufferSize - 1)

        Dim srcPtr
        Dim dstPtr

        Dim bfinal
        Dim dLen
        Dim dNLen

        Dim idx
        idx = 0

        Dim temp

        Dim i
        For i = 0 To quotient - 1
            bfinal = 0
            ret(idx) = CByte(bfinal Or (btype * 2 ^ 1))
            idx = idx + 1

            dLen = &HFFFF&
            temp = BitConverter.GetBytes(CLng(dLen), False)
            ret(idx + 0) = CByte(temp(0))
            ret(idx + 1) = CByte(temp(1))
            idx = idx + 2

            dNLen = dLen Xor &HFFFF&
            temp = BitConverter.GetBytes(dNLen, False)
            ret(idx + 0) = CByte(temp(0))
            ret(idx + 1) = CByte(temp(1))
            idx = idx + 2

            Dim j
            For j = 0 To &HFFFF& - 1
                ret(idx) = CByte(data(&HFFFF& * i + j))
                idx = idx + 1
            Next
        Next

        If remainder > 0 Then
            bfinal = CByte(1)
            ret(idx) = CByte(bfinal Or (btype * 2 ^ 1))
            idx = idx + 1

            dLen = remainder
            temp = BitConverter.GetBytes(CLng(dLen), False)
            ret(idx + 0) = CByte(temp(0))
            ret(idx + 1) = CByte(temp(1))
            idx = idx + 2

            dNLen = dLen Xor &HFFFF&
            temp = BitConverter.GetBytes(CLng(dNLen), False)
            ret(idx + 0) = CByte(temp(0))
            ret(idx + 1) = CByte(temp(1))
            idx = idx + 2

            For j = 0 To remainder - 1
                ret(idx) = CByte(data(&HFFFF& * quotient + j))
                idx = idx + 1
            Next
        End If

        Compress = ret
    End Function
End Class


Class PNG_
    Public Function GetPNG(ByRef data, _
                           ByVal pictWidth, _
                           ByVal pictHeight, _
                           ByVal foreColorRgb, _
                           ByVal backColorRgb, _
                           ByVal tColor)
        Dim bitDepth
        Select Case tColor
            Case PngColorType_pTrueColor, PngColorType_pTrueColorAlpha
                bitDepth = 8
            Case PngColorType_pIndexColor
                bitDepth = 1
            Case Else
                Call Err.Raise(5)
        End Select

        Dim psgn
        Set psgn = MakePngSignature()

        Dim ihdr
        Set ihdr = MakeIHDR( _
            pictWidth, _
            pictHeight, _
            bitDepth, _
            tColor, _
            PngCompressionMethod_Deflate, _
            PngFilterType_pNone, _
            PngInterlaceMethod_pNo _
        )

        Dim iplt
        If tColor = PngColorType_pIndexColor Then
            Set iplt = MakeIPLT(Array(foreColorRgb, backColorRgb))
        Else
            Set iplt = Nothing
        End If

        Dim idat
        Set idat = MakeIDAT(data, DeflateBType_NoCompression)

        Dim iend
        Set iend = MakeIEND()

        Dim ret
        ret = ToBytes(psgn, ihdr, iplt, idat, iend)

        Dim bw
        Set bw = New BinaryWriter

        Dim i
        For i = 0 To UBound(ret)
            bw.Append(ret(i))
        Next

        Set GetPNG = bw
    End Function

    Private Function MakePngSignature()
        Dim ret
        Set ret = New PngSignature

        With ret
            .psData(0) = CByte(&H89)
            .psData(1) = CByte(Asc("P"))
            .psData(2) = CByte(Asc("N"))
            .psData(3) = CByte(Asc("G"))
            .psData(4) = CByte(Asc(vbCr))
            .psData(5) = CByte(Asc(vbLf))
            .psData(6) = CByte(&H1A)
            .psData(7) = CByte(Asc(vbLf))
        End With

        Set MakePngSignature = ret
    End Function

    Private Function MakeIHDR(ByVal pictWidth, _
                              ByVal pictHeight, _
                              ByVal bitDepth, _
                              ByVal tColor, _
                              ByVal compression, _
                              ByVal tFilter, _
                              ByVal interlace)
        Const STR_IHDR = &H49484452

        Dim ret
        Set ret = New PngChunk

        Dim lbe
        Dim crc

        Dim temp
        Dim idx
        idx = 0

        With ret
            .pLength = 13
            .pType = STR_IHDR

            .ResizeData(.pLength - 1)
            temp = BitConverter.GetBytes(CLng(pictWidth), True)
            .pData(idx + 0) = CByte(temp(0))
            .pData(idx + 1) = CByte(temp(1))
            .pData(idx + 2) = CByte(temp(2))
            .pData(idx + 3) = CByte(temp(3))
            idx = idx + 4

            temp = BitConverter.GetBytes(CLng(pictHeight), True)
            .pData(idx + 0) = CByte(temp(0))
            .pData(idx + 1) = CByte(temp(1))
            .pData(idx + 2) = CByte(temp(2))
            .pData(idx + 3) = CByte(temp(3))
            idx = idx + 4

            .pData(idx + 0) = CByte(bitDepth)
            .pData(idx + 1) = CByte(tColor)
            .pData(idx + 2) = CByte(compression)
            .pData(idx + 3) = CByte(tFilter)
            .pData(idx + 4) = CByte(interlace)

            crc = CRC32.Checksum(BitConverter.GetBytes(STR_IHDR, True))
            .pCRC = CRC32.Update(crc, .pData)
        End With

        Set MakeIHDR = ret
    End Function

    Private Function MakeIPLT(ByRef rgbArray())
        Const STR_PLTE = &H504C5445

        Dim idx
        idx = 0

        Dim ret
        Set ret = New PngChunk

        Dim v
        Dim crc

        With ret
            .pLength = (UBound(rgbArray) + 1) * 3
            .pType = STR_PLTE

            .ResizeData(.pLength - 1)
            For Each v In rgbArray
                .pData(idx + 0) = CByte(v And &HFF&)
                .pData(idx + 1) = CByte((v And &HFF00&) \ 2 ^ 8)
                .pData(idx + 2) = CByte((v And &HFF0000) \ 2 ^ 16)
                idx = idx + 3
            Next

            crc = CRC32.Checksum(BitConverter.GetBytes(.pType, True))
            .pCRC = CRC32.Update(crc, .pData)
        End With

        Set MakeIPLT = ret
    End Function

    Private Function MakeIDAT(ByRef data, ByVal btype)
        Const STR_IDAT = &H49444154

        Dim ret
        Set ret = New PngChunk

        Dim crc

        With ret
            .pData = ZLIB.Compress(data, btype)
            .pLength = UBound(.pData) + 1
            .pType = STR_IDAT
            crc = CRC32.Checksum(BitConverter.GetBytes(STR_IDAT, True))
            .pCRC = CRC32.Update(crc, .pData)
        End With

        Set MakeIDAT = ret
    End Function

    Private Function MakeIEND()
        Const STR_IEND = &H49454E44

        Dim ret
        Set ret = New PngChunk

        With ret
            .pLength = 0
            .pType = STR_IEND
            .pCRC = CRC32.Checksum(BitConverter.GetBytes(STR_IEND, True))
        End With

        Set MakeIEND = ret
    End Function

    Private Function ToBytes(ByVal psgn, ByVal ihdr, ByVal iplt, ByVal idat, ByVal iend)
        Dim pfhSize
        pfhSize = 8

        Dim ihdrSize
        ihdrSize = 12 + ihdr.pLength

        Dim ipltSize

        If Not (iplt Is Nothing) Then
            ipltSize = 12 + iplt.pLength
        Else
            ipltSize = 0
        End If

        Dim idatSize
        idatSize = 12 + idat.pLength

        Dim iendSize
        iendSize = 12 + iend.pLength

        Dim ret
        ReDim ret(pfhSize + ihdrSize + ipltSize + idatSize + iendSize - 1)

        Dim idx
        idx = 0
        Dim i

        With psgn
            For i = 0 To UBound(.psData)
                ret(idx) = CByte(.psData(i))
                idx = idx + 1
            Next
        End With

        Dim lbe

        Dim temp
        With ihdr
            temp = ihdr.GetBytes()
            For i = 0 To UBound(temp)
                ret(idx) = CByte(temp(i))
                idx = idx + 1
            Next
        End With

        If Not (iplt Is Nothing) Then
            With iplt
                temp = iplt.GetBytes()
                For i = 0 To UBound(temp)
                    ret(idx) = CByte(temp(i))
                    idx = idx + 1
                Next
            End With
        End If

        With idat
            temp = idat.GetBytes()
            For i = 0 To UBound(temp)
                ret(idx) = CByte(temp(i))
                idx = idx + 1
            Next
        End With

        With iend
            temp = iend.GetBytes()
            For i = 0 To UBound(temp)
                ret(idx) = CByte(temp(i))
                idx = idx + 1
            Next
        End With

        ToBytes = ret
    End Function
End Class


Class ZLIB_
    Public Function Compress(ByRef data, ByVal btype)
        Dim cmf
        cmf = CByte(&H78)

        Dim fdict
        fdict = CByte(&H0)

        Dim flevel
        flevel = CByte(CompressionLevel_Default * 2 ^ 6)

        Dim flg
        flg = CByte(flevel + fdict)

        Dim fcheck
        fcheck = CByte(31 - ((cmf * 2 ^ 8 + flg) Mod 31))

        flg = flg + fcheck

        Dim bytes
        bytes = Deflate.Compress(data, btype)

        Dim adler
        adler = ADLER32.Checksum(data)

        Dim sz
        sz = UBound(bytes) + 1

        Dim ret
        ReDim ret(1 + 1 + sz + 4 - 1)

        Dim idx
        idx = 0

        Dim lbe
        ret(idx) = CByte(cmf)
        idx = idx + 1

        ret(idx) = CByte(flg)
        idx = idx + 1

        Dim i
        For i = 0 To UBound(bytes)
            ret(idx) = CByte(bytes(i))
            idx = idx + 1
        Next

        Dim temp
        temp = BitConverter.GetBytes(CLng(adler), True)
        ret(idx + 0) = CByte(temp(0))
        ret(idx + 1) = CByte(temp(1))
        ret(idx + 2) = CByte(temp(2))
        ret(idx + 3) = CByte(temp(3))

        Compress = ret
    End Function
End Class


Class PngSignature
    Public psData(7)
End Class


Class PngChunk
    Public pLength
    Public pType
    Public pData
    Public pCRC

    Private Sub Class_Initialize()
        pData = Array()
    End Sub

    Public Sub ResizeData(ByVal sz)
        ReDim pData(sz)
    End Sub

    Public Function GetBytes()
        Dim ret
        If pLength > 0 Then
            ReDim ret(12 + (UBound(pData) + 1) - 1)
        Else
            ReDim ret(12 - 1)
        End If

        Dim idx
        idx = 0

        Dim temp
        temp = BitConverter.GetBytes(CLng(pLength), True)
        ret(idx + 0) = CByte(temp(0))
        ret(idx + 1) = CByte(temp(1))
        ret(idx + 2) = CByte(temp(2))
        ret(idx + 3) = CByte(temp(3))
        idx = idx + 4

        temp = BitConverter.GetBytes(CLng(pType), True)
        ret(idx + 0) = CByte(temp(0))
        ret(idx + 1) = CByte(temp(1))
        ret(idx + 2) = CByte(temp(2))
        ret(idx + 3) = CByte(temp(3))
        idx = idx + 4

        If pLength > 0 Then
            Dim i
            For i = 0 To UBound(pData)
                ret(idx) = CByte(pData(i))
                idx = idx + 1
            Next
        End If

        temp = BitConverter.GetBytes(CLng(pCRC), True)
        ret(idx + 0) = CByte(temp(0))
        ret(idx + 1) = CByte(temp(1))
        ret(idx + 2) = CByte(temp(2))
        ret(idx + 3) = CByte(temp(3))

        GetBytes = ret
    End Function
End Class


Class BitConverter_
    Public Function GetBytes(ByVal arg, ByVal reverse)
        Dim ret
        Dim temp

        Select Case VarType(arg)
            Case vbByte
                ReDim ret(0)
                ret(0) = CByte(arg)
            Case vbInteger
                ReDim ret(1)
                ret(0) = CByte(arg And &HFF&)
                ret(1) = CByte((arg And &HFF00&) \ 2 ^ 8)

                If reverse Then
                    temp = ret(0)
                    ret(0) = CByte(ret(1))
                    ret(1) = CByte(temp)
                End If
            Case vbLong
                ReDim ret(3)
                ret(0) = CByte(arg And &HFF&)
                ret(1) = CByte((arg And &HFF00&) \ 2 ^ 8)
                ret(2) = CByte((arg And &HFF0000) \ 2 ^ 16)
                ret(3) = CByte((arg And &HFF000000) \ 2 ^ 24 And &HFF&)

                If reverse Then
                    temp = ret(0)
                    ret(0) = CByte(ret(3))
                    ret(3) = CByte(temp)

                    temp = ret(1)
                    ret(1) = CByte(ret(2))
                    ret(2) = CByte(temp)
                End If
            Case Else
                Call Err.Raise(5)
        End Select

        GetBytes = ret
    End Function
End Class
