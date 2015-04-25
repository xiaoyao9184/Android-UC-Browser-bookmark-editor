Attribute VB_Name = "Module1"
Option Explicit

Public Type ONE
    MyID As String
    FatherID As String
    Tyte As Byte
    CName As String
    WWW As String
End Type

Public iAndroid(6) As Byte 'android的UTF8编码 616E64726F6964
'Public Tyte_bk As Byte  '04
'Public Tyte_dr As Byte  '05

Public SizeL As Byte
Public Size() As Byte

Public Hmany As Long '记录有多少个书签
Public LID As String '最终ID标识
Public iALL() As ONE '记录每个书签的内容


Public nownum As Integer '目前元素NUM
Public iButton(2) As Integer '隐藏菜单
Public SaveP As String '隐藏菜单

Public F As Long  '编号

Public Declare Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hwndowner As Long, ByVal nfolder As Integer, ppidl As Long) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal szPath As String) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long



Public Function contrast(the1() As Byte, the2() As Byte, thelong) As Boolean
Dim F$, T$, i
For i = 0 To thelong
    F = F & Hex(the1(i))
    T = T & Hex(the2(i))
Next i
If F = T Then
    contrast = True
Else
    contrast = False
End If
End Function


Public Function Save3(iPath As String) As Boolean
Save3 = False
Open iPath For Binary As #1
    frmMain.StatusBar1.Panels.Item(1).Text = "保存头部"
    Dim CNameWWWlong(1) As Byte, i As Long, s As String, n As Integer, Size() As Byte
    
    Put #1, , "android"
    Put #1, , 0
    
    s = CStr(Hex(Hmany))
    s = StrTo4Str(s)
    CNameWWWlong(0) = CByte("&H" & Left(s, 2))
    CNameWWWlong(1) = CByte("&H" & Right(s, 2))
    Put #1, , CNameWWWlong
    
    CNameWWWlong(0) = CByte("&H" & Left(LID, 2))
    CNameWWWlong(1) = CByte("&H" & Right(LID, 2))
    Put #1, , CNameWWWlong

For i = 1 To Hmany
    CNameWWWlong(0) = CByte("&H" & Left(iALL(i).MyID, 2))
    CNameWWWlong(1) = CByte("&H" & Right(iALL(i).MyID, 2))
    Put #1, , CNameWWWlong
    CNameWWWlong(0) = CByte("&H" & Left(iALL(i).FatherID, 2))
    CNameWWWlong(1) = CByte("&H" & Right(iALL(i).FatherID, 2))
    Put #1, , CNameWWWlong
    Put #1, , iALL(i).Tyte
    
    s = CStr(Hex(TEXT2UTF8LONG(iALL(i).CName)))
    s = StrTo4Str(s)
    CNameWWWlong(0) = CByte("&H" & Left(s, 2))
    CNameWWWlong(1) = CByte("&H" & Right(s, 2))
    Put #1, , CNameWWWlong
    
    n = TEXT2UTF8LONG(iALL(i).CName)
    ReDim Size(n)
    Size = TEXT2UTF8(iALL(i).CName)
    Put #1, , Size
    
    If iALL(i).Tyte <> 5 Then
        s = CStr(Hex(TEXT2UTF8LONG(iALL(i).WWW)))
        s = StrTo4Str(s)
        CNameWWWlong(0) = CByte("&H" & Left(s, 2))
        CNameWWWlong(1) = CByte("&H" & Right(s, 2))
        Put #1, , CNameWWWlong
        n = TEXT2UTF8LONG(iALL(i).WWW)
        ReDim Size(n)
        Size = TEXT2UTF8(iALL(i).WWW)
        Put #1, , Size
    Else
        Put #1, , 0
    End If
    frmMain.StatusBar1.Panels.Item(1).Text = i * 100 \ Hmany & "%(" & i & "/" & Hmany & ")"
Next

frmMain.StatusBar1.Panels.Item(1).Text = "将大小写入头部"
ReDim Size(LOF(1) - 7 - 1)
Get #1, 8, Size
Seek #1, 8
        s = CStr(LOF(1) - 7)
        Put #1, , CByte(Len(s))
        Dim j() As Byte
        j = TEXT2UTF8(s)
        Put #1, , j
        Put #1, , Size
        
Close
Save3 = True
End Function




Public Function Open3(iPath As String) As Boolean
    frmMain.StatusBar1.Panels.Item(1).Text = "打开文件"
    Open3 = False
    Dim i As Long, n As Byte
    Dim CNameWWWlong(1) As Byte '缓存
    Dim CNameWWW() As Byte

    Open iPath For Binary As #1
        frmMain.StatusBar1.Panels.Item(1).Text = "读取头部"
        Get #1, , iAndroid()
        If UTF2GB(iAndroid()) <> "android" Then GoTo Fileeorr:
        Get #1, , SizeL
        ReDim Size(SizeL - 1) As Byte
        Get #1, , Size()
        Seek #1, Seek(1) + 2
        Get #1, , CNameWWWlong
        Hmany = CNameWWWlong(0) * 256 + CNameWWWlong(1)
        
        Get #1, , CNameWWWlong
        LID = CStr(Hex(CLng(CNameWWWlong(0)) * 256 + CNameWWWlong(1))) '最终ID标识
        LID = StrTo4Str(LID)
    
    ReDim iALL(Hmany)
    For i = 1 To Hmany
        frmMain.StatusBar1.Panels.Item(1).Text = i * 100 \ Hmany & "%(" & i & "/" & Hmany & ")"
        'ID标识
        Get #1, , CNameWWWlong
        iALL(i).MyID = CStr(Hex(CLng(CNameWWWlong(0)) * 256 + CNameWWWlong(1)))

        'iALL(0).MyID = IIf(Len(CStr(Hex(CNameWWWlong(1)))) < 2, "0" & CStr(Hex(CNameWWWlong(1))), CStr(Hex(CNameWWWlong(1))))
        'iALL(i).MyID = CStr(Hex(CNameWWWlong(0))) & iALL(0).MyID
        'iALL(0).MyID = ""
        '父级ID标识
        Get #1, , CNameWWWlong
        iALL(i).FatherID = CStr(Hex(CLng(CNameWWWlong(0)) * 256 + CNameWWWlong(1)))
        'iALL(i).FatherID = CStr(Hex(CNameWWWlong(0))) & CStr(Hex(CNameWWWlong(1)))
        '整理标识
        iALL(i).MyID = StrTo4Str(iALL(i).MyID)
        iALL(i).FatherID = StrTo4Str(iALL(i).FatherID)

        '类型
        Get #1, , iALL(i).Tyte
        '名称
        Get #1, , CNameWWWlong
        ReDim CNameWWW(CNameWWWlong(0) * 256 + CNameWWWlong(1) - 1) As Byte
        Get #1, , CNameWWW()
        iALL(i).CName = UTF2GB(CNameWWW())

        
        If iALL(i).Tyte <> 5 Then '不是目录
            Get #1, , CNameWWWlong
            ReDim CNameWWW(CNameWWWlong(0) * 256 + CNameWWWlong(1) - 1) As Byte
            Get #1, , CNameWWW()
            iALL(i).WWW = UTF2GB(CNameWWW())
        Else
            Seek #1, Seek(1) + 2
        End If
    Next
    '验证大小
    Dim j As String
    For i = 0 To SizeL - 1
        j = j & Chr(Size(i))
    Next
    
    If LOF(1) - 8 - SizeL <> CLng(j) Then
Fileeorr:
        MsgBox "文件格式或大小错误，无法完成打开工作！", vbOKOnly, "警告"
        frmMain.StatusBar1.Panels.Item(1).Text = "文件格式或大小错误！"
        '文件错误
        Close
        LID = "0": ReDim iALL(1):  SizeL = 0: ReDim Size(1)
        Exit Function
    End If
    Close
    Open3 = True
    frmMain.StatusBar1.Panels.Item(1).Text = "打开完成！"
    frmMain.StatusBar1.Panels.Item(2).Text = "总共" & Hmany & "个书签，占用" & j & "个字节，文件共" & j + 8 + SizeL & "个字节"
End Function

Public Function Add3() As Boolean
'创造节点
Dim i As Long
frmMain.TreeView1.Nodes.Clear
'创造跟目录（ID加了一个A后缀）
For i = 1 To UBound(iALL)
    frmMain.StatusBar1.Panels.Item(1).Text = i * 100 \ Hmany & "%(" & i & "/" & Hmany & ")"
    If iALL(i).FatherID = "FFFF" Then
        frmMain.TreeView1.Nodes.Add , 3, iALL(i).MyID & "A", iALL(i).CName
    Else
        'frmMain.TreeView1.Nodes.Add iALL(i).FatherID & "A", 4, iALL(i).MyID & "A", iALL(i).CName
    End If
Next
'创造含父级的
For i = 1 To UBound(iALL)
    frmMain.StatusBar1.Panels.Item(1).Text = i * 100 \ Hmany & "%(" & i & "/" & Hmany & ")"
    If iALL(i).FatherID = "FFFF" Then
        'frmMain.TreeView1.Nodes.Add , 3, iALL(i).MyID & "A", iALL(i).CName
    Else
    '要创建的书签的父级还没创建？
        frmMain.TreeView1.Nodes.Add iALL(i).FatherID & "A", 4, iALL(i).MyID & "A", iALL(i).CName
    End If
Next

'整理顺序
    Dim j As Integer, s As String, n As Integer, m As Integer
    '统计ROOT下有多少个含有子集的目录，放到J
    frmMain.TreeView1.SelectedItem = frmMain.TreeView1.Nodes.Item(1).FirstSibling
    Do Until frmMain.TreeView1.SelectedItem Is Nothing
        If frmMain.TreeView1.SelectedItem.Children <> 0 Then j = j + 1
        frmMain.TreeView1.SelectedItem = frmMain.TreeView1.SelectedItem.Next
    Loop
    '归位到第一个
    nownum = Get_Index(frmMain.TreeView1.Nodes.Item(1).FirstSibling.Key)
    frmMain.TreeView1.SelectedItem = frmMain.TreeView1.Nodes.Item(1).FirstSibling
    '循环J次，每次将目录提到最上面
    For n = 1 To j
        If n <> 1 Then frmMain.TreeView1.SelectedItem = frmMain.TreeView1.SelectedItem.Next
        '查找目录
        Do Until frmMain.TreeView1.SelectedItem.Children <> 0
            frmMain.TreeView1.SelectedItem = frmMain.TreeView1.SelectedItem.Next
        Loop
        nownum = Get_Index(frmMain.TreeView1.SelectedItem.Key) '上移动时用到
        '根据前者的情况上移
        Do Until frmMain.TreeView1.SelectedItem.Previous Is Nothing
            If frmMain.TreeView1.SelectedItem.Previous.Children = 0 Then
                Call frmMain.mUp_Click
            Else
                Exit Do
            End If
        Loop
        frmMain.StatusBar1.Panels.Item(1).Text = n * 100 \ j & "%(" & n & "/" & j & ")"
    Next
'刷新
frmMain.StatusBar1.Panels.Item(1).Text = "完成..."
nownum = Get_Index(frmMain.TreeView1.Nodes.Item(1).FirstSibling.Key)
Call Fash
End Function






'根据名字KEY获得数组编号
Public Function Get_Index(name As String) As Integer
Dim i As Integer
For i = 1 To UBound(iALL)
    If Left(iALL(i).MyID, 4) = Left(name, 4) Then Get_Index = i: Exit Function
Next
End Function

Public Function Out(T As Byte)

Dim a As String, b As String
Dim n As String, m As Integer
Select Case T
Case 0
    a = "UC书签|*.txt|全部文件|*.*"
Case 1
    a = "UC书签|*.html|全部文件|*.*"
Case 2
    a = "UC书签|*.*|全部文件|*.*"
    frmMain.CommonDialog1.DialogTitle = "选择一个目录进入，输入新目录名保存即可"
Case 3
    a = "UC书签|*.txt|全部文件|*.*"
    b = iALL(nownum).CName
Case 4
    a = "UC书签|*.html|全部文件|*.*"
    b = iALL(nownum).CName
Case 5
    a = "UC书签|*.url|全部文件|*.*"
    b = iALL(nownum).CName
End Select



frmMain.CommonDialog1.Filter = a
frmMain.CommonDialog1.FileName = b
frmMain.CommonDialog1.ShowSave

If frmMain.CommonDialog1.FileName = "" Then Exit Function

Dim i As Integer
Open frmMain.CommonDialog1.FileName For Output As #1

If T = 0 Then '全部TXT

    For i = 1 To UBound(iALL)
        Print #1, iALL(i).CName
        Print #1, iALL(i).WWW
        Print #1, Chr(13) '& Chr(10)
    Next
    Close #1

ElseIf T = 1 Then '全部HTLM
    Dim j As String
    For i = 1 To UBound(iALL)
        Print #1, "<p>" & iALL(i).CName & "</p>"
        Print #1, "<p><a href=" & Chr(34) & iALL(i).WWW & Chr(34) & " Target = " & Chr(34) & "_blank" & Chr(34) & "class=" & Chr(34) & "weblink" & Chr(34) & ">" & iALL(i).WWW & "</a></p>"
        Print #1, "<p>&nbsp;</p>"
    Next
    Close #1

ElseIf T = 2 Then '全部URL
    Close
    Kill frmMain.CommonDialog1.FileName
    Dim F As String, iname As String
    
    F = frmMain.CommonDialog1.FileName
    If Dir(Replace(F & "\", "\\", "\")) = "" Then MkDir Replace(F, "\\", "\")
    ChDir Replace(F & "\", "\\", "\")
    For i = 1 To UBound(iALL) '查找数据
        If iALL(i).Tyte <> 5 Then
            If iALL(i).FatherID = "FFFF" Then '判断有无父级
                iname = Replace(F & "\" & iALL(i).CName & ".url", "\\", "\")
                iname = Replace(iname, "/", ".")
                iname = Replace(iname, "*", ".")
                iname = Replace(iname, "?", ".")
                iname = Replace(iname, "|", ".")
                iname = Replace(Replace(iname, "<", "."), ">", ".")
                Open iname For Output As #1
                    Print #1, "[InternetShortcut]"
                    Print #1, "URL=" & iALL(i).WWW
                Close #1
            Else
                j = 0
                Do '查找父级名称
                    j = j + 1
                Loop Until iALL(i).FatherID = iALL(j).MyID
                F = iALL(j).CName
                If Dir(Replace(F & "\", "\\", "\")) = "" Then MkDir Replace(F, "\\", "\")
                'On Error Resume Next
                On Error Resume Next
                iname = Replace(F & "\" & iALL(i).CName & ".url", "\\", "\")
                iname = Replace(iname, "/", ".")
                iname = Replace(iname, "*", ".")
                iname = Replace(iname, "?", ".")
                iname = Replace(iname, "|", ".")
                iname = Replace(Replace(iname, "<", "."), ">", ".")
                
                Open iname For Output As #1
                    Print #1, "[InternetShortcut]"
                    Print #1, "URL=" & iALL(i).WWW
                Close #1
            End If

        Else
            If Dir(Replace(F & "\" & iALL(i).CName, "\\", "\")) <> "" Then MkDir Replace(F & "\" & iALL(i).CName, "\\", "\")
        End If

    Next
ElseIf T = 3 Then '一个TXT
    If frmMain.TreeView1.Nodes(iALL(nownum).MyID & "A").Children Then
        n = Left(frmMain.TreeView1.Nodes(iALL(nownum).MyID & "A").Child.Key, 4) '取得该级下的第一个子集
        
        For m = 1 To frmMain.TreeView1.Nodes(iALL(nownum).MyID & "A").Children
            For i = 1 To UBound(iALL) '查找数据
                If n = iALL(i).MyID Then
                    Print #1, iALL(i).CName
                    Print #1, iALL(i).WWW
                    Print #1, Chr(13) '& Chr(10)
                End If
            Next
            If Not (frmMain.TreeView1.Nodes(n & "A").Next Is Nothing) Then n = Left(frmMain.TreeView1.Nodes(n & "A").Next.Key, 4) '下一个子集
        Next
        'For i = 1 To UBound(iALL)
        '    If iALL(i).FatherID = iALL(nownum).MyID Then
        '        Print #1, iALL(i).CName
        '        Print #1, iALL(i).WWW
        '        Print #1, Chr(13) '& Chr(10)
        '    End If
        'Next
    Else
        Print #1, iALL(nownum).CName
        Print #1, iALL(nownum).WWW
        Print #1, Chr(13) '& Chr(10)
    End If
ElseIf T = 4 Then '一个HTLM
    If frmMain.TreeView1.Nodes(iALL(nownum).MyID & "A").Children Then
        n = Left(frmMain.TreeView1.Nodes(iALL(nownum).MyID & "A").Child.Key, 4) '取得该级下的第一个子集
        
        For m = 1 To frmMain.TreeView1.Nodes(iALL(nownum).MyID & "A").Children
            For i = 1 To UBound(iALL) '查找数据
                If n = iALL(i).MyID Then
                    Print #1, "<p>" & iALL(i).CName & "</p>"
                    Print #1, "<p><a href=" & Chr(34) & iALL(i).WWW & Chr(34) & " Target = " & Chr(34) & "_blank" & Chr(34) & "class=" & Chr(34) & "weblink" & Chr(34) & ">" & iALL(i).WWW & "</a></p>"
                    Print #1, "<p>&nbsp;</p>"
                End If
            Next
            If Not (frmMain.TreeView1.Nodes(n & "A").Next Is Nothing) Then n = Left(frmMain.TreeView1.Nodes(n & "A").Next.Key, 4) '下一个子集
        Next
        
        'For i = 1 To UBound(iALL)
        '    If iALL(i).FatherID = iALL(nownum).MyID Then
        '        Print #1, "<p>" & iALL(i).CName & "</p>"
        '        Print #1, "<p><a href=" & Chr(34) & iALL(i).WWW & Chr(34) & " Target = " & Chr(34) & "_blank" & Chr(34) & "class=" & Chr(34) & "weblink" & Chr(34) & ">" & iALL(i).WWW & "</a></p>"
        '        Print #1, "<p>&nbsp;</p>"
        '    End If
        'Next
    Else
        Print #1, "<p>" & iALL(nownum).CName & "</p>"
        Print #1, "<p><a href=" & Chr(34) & iALL(nownum).WWW & Chr(34) & " Target = " & Chr(34) & "_blank" & Chr(34) & "class=" & Chr(34) & "weblink" & Chr(34) & ">" & iALL(nownum).WWW & "</a></p>"
        Print #1, "<p>&nbsp;</p>"
    End If
ElseIf T = 5 Then '一个URL
    If frmMain.TreeView1.Nodes(iALL(nownum).MyID & "A").Children Then
        Close
        Kill frmMain.CommonDialog1.FileName
        MkDir Left(frmMain.CommonDialog1.FileName, Len(frmMain.CommonDialog1.FileName) - 4)
            n = Left(frmMain.TreeView1.Nodes(iALL(nownum).MyID & "A").Child.Key, 4) '取得该级下的第一个子集
            For m = 1 To frmMain.TreeView1.Nodes(iALL(nownum).MyID & "A").Children
                For i = 1 To UBound(iALL) '查找数据
                    If n = iALL(i).MyID Then
                        Open Left(frmMain.CommonDialog1.FileName, Len(frmMain.CommonDialog1.FileName) - 4) & "\" & iALL(i).CName & ".url" For Output As #1
                            Print #1, "[InternetShortcut]"
                            Print #1, "URL=" & iALL(nownum).WWW
                        Close #1
                    End If
                Next
                If Not (frmMain.TreeView1.Nodes(n & "A").Next Is Nothing) Then n = Left(frmMain.TreeView1.Nodes(n & "A").Next.Key, 4) '下一个子集
            Next
    Else
        Print #1, "[InternetShortcut]"
        Print #1, "URL=" & iALL(nownum).WWW
    End If
End If
 Close
End Function



'整理数组顺序，是一个套嵌函数，以SelectedItem属性开始，四个出口
Public Function PutOneToOne()
Dim i1 As Long, n As Byte, s As String

Do Until frmMain.TreeView1.SelectedItem Is Nothing Or F > Hmany
    frmMain.StatusBar1.Panels.Item(1).Text = F * 100 \ Hmany & "%(" & F & "/" & Hmany & ")"
    '变换当前元素的在数字内的编号
    nownum = Get_Index(frmMain.TreeView1.SelectedItem.Key)
    iALL(0) = iALL(F)
    iALL(F) = iALL(nownum)
    iALL(nownum) = iALL(0)
        'F元素的父级ID，是现在选中的那个的父级，在iAll中的顺序标号
    If Not (frmMain.TreeView1.SelectedItem.Parent Is Nothing) Then '存在父级
        '取得选中项的父级KEY，在iAll中找到其标号，将标号转换为字符串
        s = CStr(Hex(Get_Index(frmMain.TreeView1.SelectedItem.Parent.Key)))
        iALL(F).FatherID = StrTo4Str(s)
    End If
    '转移到下一个目标
    If frmMain.TreeView1.SelectedItem.Children <> 0 Then '存在子级'出口1
        If Not (frmMain.TreeView1.SelectedItem.Parent Is Nothing) Then MsgBox "这是一个含有套嵌目录的UC书签，UC浏览器不支持这种格式。软件将退出", vbOKOnly, "格式错误": End
        F = F + 1
        frmMain.TreeView1.SelectedItem = frmMain.TreeView1.SelectedItem.Child '会自动打开折叠
        Call PutOneToOne '套嵌
    ElseIf Not (frmMain.TreeView1.SelectedItem.Next Is Nothing) Then  '存在下一个'出口2
        F = F + 1
        frmMain.TreeView1.SelectedItem = frmMain.TreeView1.SelectedItem.Next
    ElseIf Not (frmMain.TreeView1.SelectedItem.Parent Is Nothing) Then '存在父级'出口3
        F = F + 1
        frmMain.TreeView1.SelectedItem = frmMain.TreeView1.SelectedItem.Parent.Next: _
        frmMain.TreeView1.SelectedItem.Previous.Expanded = False: _
        Exit Do '折叠：这两行不能换，若先折叠，SelectedItem就消失
    Else '到达最后一项
        Dim j As Long
        '将iALL中的所有元素的MyID，变成其在iALL中的顺序标号
        For j = 1 To UBound(iALL)
            '将KEY全部，变成其在iALL中的顺序标号+B
            frmMain.TreeView1.Nodes.Item(iALL(j).MyID & "A").Key = StrTo4Str(CStr(Hex(j))) & "B"
            iALL(j).MyID = StrTo4Str(CStr(Hex(j)))
        Next
        '将KEY全部，变成其在iALL中的顺序标号
        For j = 1 To UBound(iALL)
            frmMain.TreeView1.Nodes.Item(j).Key = Left(frmMain.TreeView1.Nodes.Item(j).Key, 4) & "A"
        Next
        '归位
        If Hmany = UBound(iALL) Then
            frmMain.StatusBar1.Panels.Item(1).Text = "完成"
            LID = iALL(Hmany).MyID
            nownum = 1 '同Get_Index(TreeView1.Nodes.Item(1).Root.Key)
            Call Fash
            Exit Function '出口4
        Else
            MsgBox "数据严重错误，请联系开发者.错误代号PutOneToOne", vbOKOnly, "严重错误！": End
        End If
        frmMain.TreeView1.SelectedItem = Nothing '这句没用
    End If
Loop
End Function
'刷新
Public Function Fash()
    Set frmMain.TreeView1.SelectedItem = frmMain.TreeView1.Nodes.Item(iALL(nownum).MyID & "A")
    frmMain.TreeView1.DropHighlight = frmMain.TreeView1.SelectedItem
    frmMain.txtName.Text = iALL(nownum).CName
    frmMain.ChkType.Value = IIf(iALL(nownum).Tyte = 5, 1, 0) '5目录
    frmMain.txtWWW.Text = IIf(iALL(nownum).Tyte = 5, "", iALL(nownum).WWW)
    frmMain.txtWWW.Enabled = IIf(iALL(nownum).Tyte = 5, False, True)
    
    frmMain.CmdU.Enabled = IIf(frmMain.TreeView1.SelectedItem.Previous Is Nothing, False, True)
    frmMain.CmdD.Enabled = IIf(frmMain.TreeView1.SelectedItem.Next Is Nothing, False, True)
    frmMain.mUp.Enabled = frmMain.CmdU.Enabled: frmMain.mDown.Enabled = frmMain.CmdD.Enabled

End Function
'将字符凑成4个字符组成
Public Function StrTo4Str(l As String) As String
    Dim n As Byte
    For n = 1 To 4 - Len(l)
        l = "0" & l
    Next
    StrTo4Str = l
End Function
