Attribute VB_Name = "TXT2UTF8"


Public Function TEXT2UTF8(txt) '将文本转换为UTF编码数组
On Error GoTo OutError
Dim dat() As Byte, dat1() As Byte, st As String
Dim z As String
Dim zTem As String
Dim zHex As String
Dim zBin As String
Dim zRes As Long
Dim zAsc As Long
Dim l As Long
Dim i1 As Long
Dim i2 As Integer

For i1 = 1 To Len(txt)
z = Mid(txt, i1, 1): zAsc = Asc(z)
If zAsc > 0 Then
    ReDim Preserve dat(l) As Byte
    dat(l) = zAsc: l = l + 1
Else
    ReDim Preserve dat(l + 2) As Byte
    dat1 = z: st = Right("0" & Hex(dat1(1)), 2) & Right("0" & Hex(dat1(0)), 2)
    For i2 = 1 To Len(st)
        z = Mid(st, i2, 1)
        Select Case z
        Case Is = "0": zBin = zBin & "0000"
        Case Is = "1": zBin = zBin & "0001"
        Case Is = "2": zBin = zBin & "0010"
        Case Is = "3": zBin = zBin & "0011"
        Case Is = "4": zBin = zBin & "0100"
        Case Is = "5": zBin = zBin & "0101"
        Case Is = "6": zBin = zBin & "0110"
        Case Is = "7": zBin = zBin & "0111"
        Case Is = "8": zBin = zBin & "1000"
        Case Is = "9": zBin = zBin & "1001"
        Case Is = "A": zBin = zBin & "1010"
        Case Is = "B": zBin = zBin & "1011"
        Case Is = "C": zBin = zBin & "1100"
        Case Is = "D": zBin = zBin & "1101"
        Case Is = "E": zBin = zBin & "1110"
        Case Is = "F": zBin = zBin & "1111"
        End Select
    Next
    zTem = "1110" & Left(zBin, 4) & "10" & Mid(zBin, 5, 6) & "10" & Right(zBin, 6)
    For i2 = 1 To 24
        z = Mid(zTem, i2, 1)
        zRes = zRes + IIf(z = "1", 1 * 2 ^ (24 - i2), 0 * 2 ^ (24 - i2))
    Next
    z = Hex(zRes)
    dat(l) = Val("&H" & Left(z, 2))
    dat(l + 1) = Val("&H" & Mid(z, 3, 2))
    dat(l + 2) = Val("&H" & Right(z, 2))
    l = l + 3
End If
zBin = "": zRes = 0
Next

OutError:
'Close
TEXT2UTF8 = dat
End Function
Public Function TEXT2UTF8LONG(txt)
On Error GoTo OutError
Dim dat() As Byte, dat1() As Byte, st As String
Dim z As String
Dim zTem As String
Dim zHex As String
Dim zBin As String
Dim zRes As Long
Dim zAsc As Long
Dim l As Long
Dim i1 As Long
Dim i2 As Integer

For i1 = 1 To Len(txt)
z = Mid(txt, i1, 1): zAsc = Asc(z)
If zAsc > 0 Then
    ReDim Preserve dat(l) As Byte
    dat(l) = zAsc: l = l + 1
Else
    ReDim Preserve dat(l + 2) As Byte
    dat1 = z: st = Right("0" & Hex(dat1(1)), 2) & Right("0" & Hex(dat1(0)), 2)
    For i2 = 1 To Len(st)
        z = Mid(st, i2, 1)
        Select Case z
        Case Is = "0": zBin = zBin & "0000"
        Case Is = "1": zBin = zBin & "0001"
        Case Is = "2": zBin = zBin & "0010"
        Case Is = "3": zBin = zBin & "0011"
        Case Is = "4": zBin = zBin & "0100"
        Case Is = "5": zBin = zBin & "0101"
        Case Is = "6": zBin = zBin & "0110"
        Case Is = "7": zBin = zBin & "0111"
        Case Is = "8": zBin = zBin & "1000"
        Case Is = "9": zBin = zBin & "1001"
        Case Is = "A": zBin = zBin & "1010"
        Case Is = "B": zBin = zBin & "1011"
        Case Is = "C": zBin = zBin & "1100"
        Case Is = "D": zBin = zBin & "1101"
        Case Is = "E": zBin = zBin & "1110"
        Case Is = "F": zBin = zBin & "1111"
        End Select
    Next
    zTem = "1110" & Left(zBin, 4) & "10" & Mid(zBin, 5, 6) & "10" & Right(zBin, 6)
    For i2 = 1 To 24
        z = Mid(zTem, i2, 1)
        zRes = zRes + IIf(z = "1", 1 * 2 ^ (24 - i2), 0 * 2 ^ (24 - i2))
    Next
    z = Hex(zRes)
    dat(l) = Val("&H" & Left(z, 2))
    dat(l + 1) = Val("&H" & Mid(z, 3, 2))
    dat(l + 2) = Val("&H" & Right(z, 2))
    l = l + 3
End If
zBin = "": zRes = 0
Next

OutError:
'Close
TEXT2UTF8LONG = l
End Function



'UTF8编码文字将转换为汉字
Function ConvChinese(ByRef X, Y As Integer)
Dim a(2)
Dim i, j, DigS, Unicode
i = 0
j = 0

For i = Y To Y + 2
    a(i - Y) = Hex(X(i))
Next

For i = 0 To UBound(a)
    a(i) = c16to2(a(i))
Next

For i = 0 To UBound(a) - 1
    DigS = InStr(a(i), "0")
    Unicode = ""
    For j = 1 To DigS - 1
        If j = 1 Then
            a(i) = Right(a(i), Len(a(i)) - DigS)
            Unicode = Unicode & a(i)
        Else
            i = i + 1
            a(i) = Right(a(i), Len(a(i)) - 2)
            Unicode = Unicode & a(i)
        End If
    Next
    If Len(c2to16(Unicode)) = 4 Then
        ConvChinese = ConvChinese & ChrW(Int("&H" & c2to16(Unicode)))
    Else
        ConvChinese = ConvChinese & Chr(Int("&H" & c2to16(Unicode)))
    End If
Next
End Function

'二进制代码转换为十六进制代码
Function c2to16(X)
Dim i
i = 1
For i = 1 To Len(X) Step 4
    c2to16 = c2to16 & Hex(c2to10(Mid(X, i, 4)))
Next
End Function

'二进制代码转换为十进制代码
Function c2to10(X)
Dim i
c2to10 = 0
If X = "0" Then Exit Function
i = 0
For i = 0 To Len(X) - 1
    If Mid(X, Len(X) - i, 1) = "1" Then c2to10 = c2to10 + 2 ^ (i)
Next
End Function

'十六进制代码转换为二进制代码
Function c16to2(X)
Dim i, tempstr
i = 0
For i = 1 To Len(Trim(X))
    tempstr = c10to2(CInt(Int("&h" & Mid(X, i, 1))))
    Do While Len(tempstr) < 4
        tempstr = "0" & tempstr
    Loop
    c16to2 = c16to2 & tempstr
Next
End Function

'十进制代码转换为二进制代码
Function c10to2(X)
Dim mysign, DigS, tempnum, i
mysign = Sgn(X)
X = Abs(X)
DigS = 1
Do
    If X < 2 ^ DigS Then
        Exit Do
    Else
        DigS = DigS + 1
    End If
Loop
tempnum = X
i = 0
For i = DigS To 1 Step -1
    If tempnum >= 2 ^ (i - 1) Then
        tempnum = tempnum - 2 ^ (i - 1)
        c10to2 = c10to2 & "1"
    Else
        c10to2 = c10to2 & "0"
    End If
Next
If mysign = -1 Then c10to2 = "-" & c10to2
End Function



'将UTF编码转换为文本
Function UTF2GB(ByRef UTFStr) 'UTFStr数组类型，输出为文本
Dim Dig As Integer
Dim GBStr As String
    For Dig = 0 To UBound(UTFStr)
        If UTFStr(Dig) = 194 Then
            'GBStr = GBStr & Chr(UTFStr(Dig + 1))
            ''''''''''''''''丢失中文符号
            Dig = Dig + 1
        ElseIf UTFStr(Dig) > 127 Then   'UTF8编码文字大于127则转换为汉字
            GBStr = GBStr & ConvChinese(UTFStr, Dig)
            Dig = Dig + 2 '如果UTF8编码汉字跳过三个字节
        Else
            GBStr = GBStr & Chr(UTFStr(Dig))
        End If
    Next
UTF2GB = GBStr
End Function
