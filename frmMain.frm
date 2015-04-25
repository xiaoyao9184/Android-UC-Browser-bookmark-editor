VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Android UC 书签编辑器"
   ClientHeight    =   5115
   ClientLeft      =   2700
   ClientTop       =   2340
   ClientWidth     =   8940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   8940
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   11
      Top             =   4785
      Width           =   8940
      _ExtentX        =   15769
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   9596
            MinWidth        =   9596
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton CmdD 
      Caption         =   "下移"
      Height          =   1935
      Left            =   120
      TabIndex        =   10
      Top             =   2280
      Width           =   255
   End
   Begin VB.CommandButton CmdU 
      Caption         =   "上移"
      Height          =   1935
      Left            =   120
      TabIndex        =   9
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton CmdDele 
      Caption         =   "删除"
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   4320
      Width           =   1815
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "添加   ->"
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   4320
      Width           =   1695
   End
   Begin VB.CommandButton CmdGO 
      Caption         =   "访问网址"
      Height          =   495
      Left            =   7080
      TabIndex        =   6
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CheckBox ChkType 
      Caption         =   "类型：书签/目录"
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   1800
      Width           =   2175
   End
   Begin VB.TextBox txtWWW 
      Height          =   495
      Left            =   4320
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   3000
      Width           =   4095
   End
   Begin VB.TextBox txtName 
      Height          =   495
      Left            =   4320
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1080
      Width           =   4095
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   3855
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   6800
      _Version        =   393217
      Indentation     =   529
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8400
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Caption         =   "网址"
      Height          =   255
      Left            =   4320
      TabIndex        =   5
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "标题"
      Height          =   255
      Left            =   4320
      TabIndex        =   4
      Top             =   720
      Width           =   1335
   End
   Begin VB.Menu mFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mNew 
         Caption         =   "新建(&N)"
      End
      Begin VB.Menu mOpen 
         Caption         =   "打开(&O)"
      End
      Begin VB.Menu mSave 
         Caption         =   "保存(&S)"
      End
      Begin VB.Menu mAsave 
         Caption         =   "另存为(&A)..."
      End
      Begin VB.Menu mExit 
         Caption         =   "退出(&E)"
      End
   End
   Begin VB.Menu mEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mAdd 
         Caption         =   "添加(&A)"
         Begin VB.Menu mFront 
            Caption         =   "添加到前面(&F)"
         End
         Begin VB.Menu mBehind 
            Caption         =   "添加到后面(&B)"
         End
         Begin VB.Menu mSub 
            Caption         =   "添加子级(&S)"
         End
      End
      Begin VB.Menu mUp 
         Caption         =   "上移(&U)"
      End
      Begin VB.Menu mDown 
         Caption         =   "下移(&D)"
      End
      Begin VB.Menu mDele 
         Caption         =   "删除(&E)"
      End
      Begin VB.Menu mf0 
         Caption         =   "-"
      End
      Begin VB.Menu mOut 
         Caption         =   "导出(&O)"
         Begin VB.Menu mTXT 
            Caption         =   "导出全部为TXT"
         End
         Begin VB.Menu mHTML 
            Caption         =   "导出全部为HTML"
         End
         Begin VB.Menu mUrl 
            Caption         =   "导出全部为URL"
         End
         Begin VB.Menu mFavorites 
            Caption         =   "导出全部到收藏夹"
         End
         Begin VB.Menu mf1 
            Caption         =   "-"
         End
         Begin VB.Menu mOneTXT 
            Caption         =   "导出选中项目为TXT"
         End
         Begin VB.Menu mOneHTML 
            Caption         =   "导出选中项目为HTML"
         End
         Begin VB.Menu mOneURL 
            Caption         =   "导出选中项目为URL"
         End
         Begin VB.Menu mOneFavorites 
            Caption         =   "导出选中项目到收藏夹"
         End
         Begin VB.Menu mf2 
            Caption         =   "-"
         End
      End
      Begin VB.Menu mIn 
         Caption         =   "导入(&O)"
         Begin VB.Menu mInFavorites 
            Caption         =   "导入收藏夹"
         End
      End
   End
   Begin VB.Menu mAbout 
      Caption         =   "关于(&A)"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL = 1

Private Sub CmdAdd_Click()
    If iButton(0) = 3 Then frmMain.PopupMenu mAdd, vbPopupMenuLeftAlign, iButton(1) + CmdAdd.Left, iButton(2) + CmdAdd.Top
End Sub
Private Sub CmdAdd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    iButton(0) = 3: iButton(1) = X: iButton(2) = Y
End Sub

Private Sub CmdDele_Click()
    Call mdele_Click
End Sub
Private Sub CmdU_Click()
    Call mUp_Click
End Sub
Private Sub CmdD_Click()
    Call mDown_Click
End Sub
Private Sub CmdGO_Click()
'访问网址
'Text2.Text = UTF2GB(Text1.Text)





'函数返回值为大于32的整数表明成功执行调用，小于或等于32表明调用失败。
Dim Result

'Result = ShellExecute(0, vbNullString, "http://bbs.angodroid.com/?fromuid=818", vbNullString, vbNullString, SW_SHOWNORMAL)
'If Result <= 32 Then
'    MsgBox "调用浏览器错误！", vbOKOnly + vbCritical, "错误：", 0
'End If
Result = ShellExecute(0, vbNullString, txtWWW.Text, vbNullString, vbNullString, SW_SHOWNORMAL)
If Result <= 32 Then
    MsgBox "调用浏览器错误！", vbOKOnly + vbCritical, "错误：", 0
End If
End Sub








Private Sub Form_Load()
    LID = "0"
    '菜单
    mSave.Enabled = False: mAsave.Enabled = False: mEdit.Enabled = False
    mIn.Enabled = False
    '按钮
    CmdAdd.Enabled = False: CmdDele.Enabled = False
    CmdU.Enabled = False: CmdD.Enabled = False
    CmdGO.Enabled = False: ChkType.Enabled = False
    txtName.Enabled = False: txtWWW.Enabled = False
    
    StatusBar1.Panels.Item(1).MinWidth = 1500
    StatusBar1.Panels.Item(2).MinWidth = frmMain.Width - StatusBar1.Panels.Item(1).MinWidth
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    frmMain.StatusBar1.Panels.Item(2).Text = ""
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Unload Me
    End
End Sub


'×××××××××××××菜单××××××××××××××
Private Sub mExit_Click()
    End
End Sub
Private Sub mAbout_Click()
    frmMain.Enabled = False
    frmAbout.Show
End Sub

Private Sub mInFavorites_Click()
    Dim sName As String, sFile As String, sDirList() As String, iDirNum As Integer
    Dim sDirName As String
    Dim i As Integer
    
Dim kk As New IWshRuntimeLibrary.IWshShell_Class
'先在工程引用添加 windows script host object model
'txtName = Environ("username") 取得当前用户名
sDirName = kk.SpecialFolders("FAVORITES")

    '首先枚举所有文件夹
    sName = Dir(sDirName + "\", vbDirectory)
    Do While Len(sName) > 0
        If sName <> "." And sName <> ".." And Right(sName, 4) <> ".url" Then
            iDirNum = iDirNum + 1
            ReDim Preserve sDirList(1 To iDirNum) '重定义
            sDirList(iDirNum) = sDirName + "\" + sName + "\"  '取得目录数组
        End If
        sName = Dir '下一个文件夹
    Loop
    '所有文件夹下的文件
    For i = 1 To UBound(sDirList)
    
        ''''''''''''''''''''''''''''''创建目录
        '用函数将标题在现有节点中扫描一遍重复的不添加
        '添加时取得最后一个ID并保存ID
        sFile = Dir(sDirList(i) + "\*.url", vbDirectory + vbNormal)
        Do While Len(sFile) > 0

        
        

            '''''''''''''''''''''''''''创建
            sFile = Dir '下一个文件
        Loop
    Next
    '所有文件
    sFile = Dir(txtName.Text + "\*.url", vbDirectory + vbNormal)   '包括隐藏文件
    Do While Len(sFile) > 0
        '''''''''''''''''''''''''''创建
        sFile = Dir '下一个文件
    Loop

End Sub



Private Sub mNew_Click()
    frmMain.TreeView1.Nodes.Clear
    Hmany = 0: LID = "0000"

    mAddSub (1)
    mAsave.Enabled = True: mEdit.Enabled = True
    CmdAdd.Enabled = True: CmdDele.Enabled = True
    txtName.Enabled = True: ChkType.Enabled = True
End Sub

Private Sub mOpen_Click()
Dim line$, i$
Dim r%, c As Byte

CommonDialog1.Filter = "UC书签|*.aucf|全部文件|*.*"
CommonDialog1.ShowOpen
If CommonDialog1.FileName = "" Or Err.Number = 32755 Then Exit Sub
If Open3(CommonDialog1.FileName) = False Then Exit Sub

    Add3
    
    F = 1
    TreeView1.SelectedItem = TreeView1.Nodes(1).Root
    Call PutOneToOne
    
    SaveP = CommonDialog1.FileName
    
    mSave.Enabled = True: mAsave.Enabled = True: mEdit.Enabled = True
    CmdAdd.Enabled = True: CmdDele.Enabled = True
    txtName.Enabled = True: ChkType.Enabled = True
    
    StatusBar1.Panels.Item(1).Text = "完成！"
End Sub
Private Sub mSave_Click()
    Kill SaveP
    F = 1
    TreeView1.SelectedItem = TreeView1.Nodes(1).Root
    Call PutOneToOne
    Save3 (SaveP)
    frmMain.StatusBar1.Panels.Item(1).Text = "完成！"
End Sub

Private Sub mAsave_Click()
CommonDialog1.Filter = "UC书签|*.aucf|全部文件|*.*"
CommonDialog1.ShowSave
If CommonDialog1.FileName = "" Or Err.Number = 32755 Then Exit Sub

    F = 1
    TreeView1.SelectedItem = TreeView1.Nodes(1).Root
    Call PutOneToOne
    
    If Save3(CommonDialog1.FileName) = False Then Exit Sub
        SaveP = CommonDialog1.FileName
        frmMain.StatusBar1.Panels.Item(1).Text = "完成！"
End Sub


Private Sub mEdit_Click()
mTXT.Enabled = True
mHTML.Enabled = True
mUrl.Enabled = True
mFavorites.Enabled = True
End Sub

Private Sub mTXT_Click()
 Out (0)
End Sub
Private Sub mHTML_Click()
 Out (1)
End Sub
Private Sub mURL_Click()
 Out (2)
End Sub
Private Sub mOneTXT_Click()
 Out (3)
End Sub
Private Sub mOneHTML_Click()
 Out (4)
End Sub
Private Sub mOneURL_Click()
 Out (5)
End Sub
Private Sub mFavorites_Click()
    Dim sDirName As String
    Dim i As Integer, j As Integer, n As String, F As String
    Dim iname As String
    Dim kk As New IWshRuntimeLibrary.IWshShell_Class
    '先在工程引用添加 windows script host object model
    'txtName = Environ("username") 取得当前用户名
    sDirName = kk.SpecialFolders("FAVORITES")


    For i = 1 To UBound(iALL) '查找数据
        If iALL(i).Tyte <> 5 Then
            If iALL(i).FatherID = "FFFF" Then '判断有无父级
                iname = Replace(sDirName & "\" & F & "\" & iALL(i).CName & ".url", "\\", "\")
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
                If Dir(Replace(sDirName & "\" & F & "\", "\\", "\")) = "" Then MkDir Replace(sDirName & "\" & F, "\\", "\")
                'On Error Resume Next
                On Error Resume Next
                iname = Replace(sDirName & "\" & F & "\" & iALL(i).CName & ".url", "\\", "\")
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
            If Dir(Replace(sDirName & "\" & iALL(i).CName, "\\", "\")) <> "" Then MkDir Replace(sDirName & "\" & iALL(i).CName, "\\", "\")
        End If

    Next
End Sub
Private Sub mOneFavorites_Click()
    Dim sDirName As String
    Dim i As Integer, n As String, m As Integer
    
    Dim kk As New IWshRuntimeLibrary.IWshShell_Class
    '先在工程引用添加 windows script host object model
    'txtName = Environ("username") 取得当前用户名
    sDirName = kk.SpecialFolders("FAVORITES")

    If frmMain.TreeView1.Nodes(iALL(nownum).MyID & "A").Children Then
        If Dir(Replace(sDirName & "\" & iALL(nownum).CName, "\\", "\")) = "" Then MkDir Replace(sDirName & "\" & iALL(nownum).CName, "\\", "\")
            n = Left(frmMain.TreeView1.Nodes(iALL(nownum).MyID & "A").Child.Key, 4) '取得该级下的第一个子集
            For m = 1 To frmMain.TreeView1.Nodes(iALL(nownum).MyID & "A").Children
                For i = 1 To UBound(iALL) '查找数据
                    If n = iALL(i).MyID Then
                        Open Replace(sDirName & "\" & iALL(nownum).CName & "\" & iALL(i).CName & ".url", "\\", "\") For Output As #1
                            Print #1, "[InternetShortcut]"
                            Print #1, "URL=" & iALL(i).WWW
                        Close #1
                    End If
                Next
                If Not (frmMain.TreeView1.Nodes(n & "A").Next Is Nothing) Then n = Left(frmMain.TreeView1.Nodes(n & "A").Next.Key, 4) '下一个子集
            Next
    Else
        Open Replace(sDirName & "\" & iALL(nownum).CName & ".url", "\\", "\") For Output As #1
            Print #1, "[InternetShortcut]"
            Print #1, "URL=" & iALL(nownum).WWW
        Close #1
    End If
End Sub

Public Sub mUp_Click() '上
    Dim ilast As String, i As Integer
    ilast = TreeView1.SelectedItem.Previous.Key
    TreeView1.SelectedItem.Key = TreeView1.SelectedItem.Key & "B"
    
    TreeView1.Nodes.Add ilast, 3, iALL(nownum).MyID & "A", iALL(nownum).CName
    '把子集移动过来
        For i = 1 To TreeView1.SelectedItem.Children
            Set TreeView1.SelectedItem.Child.Parent = TreeView1.Nodes.Item(iALL(nownum).MyID & "A")
        Next
        TreeView1.Nodes.Item(iALL(nownum).MyID & "A").Expanded = False '关闭折叠
    TreeView1.Nodes.Remove (TreeView1.SelectedItem.Index)
    Call Fash
End Sub
Public Sub mDown_Click() '下
    Dim inext As String, i As Integer
    inext = TreeView1.SelectedItem.Next.Key
    TreeView1.SelectedItem.Key = TreeView1.SelectedItem.Key & "B"

    TreeView1.Nodes.Add inext, 2, iALL(nownum).MyID & "A", iALL(nownum).CName
    '把子集移动过来
        For i = 1 To TreeView1.SelectedItem.Children
            Set TreeView1.SelectedItem.Child.Parent = TreeView1.Nodes.Item(iALL(nownum).MyID & "A")
        Next
        TreeView1.Nodes.Item(iALL(nownum).MyID & "A").Expanded = False '关闭折叠
    TreeView1.Nodes.Remove (TreeView1.SelectedItem.Index)
    Call Fash
End Sub
Private Sub mdele_Click() '删除
    '有子级书签，删除子级数组，不包括本身
    If TreeView1.SelectedItem.Children <> 0 Then
        Dim Fk As String, n As Long, i As Integer
        Fk = Left(TreeView1.SelectedItem.Key, 4) '选中的KEY
        
        For n = 1 To TreeView1.SelectedItem.Children
            For i = 1 To UBound(iALL) '挑选是Fk的
                If Left(iALL(i).FatherID, 4) = Fk Then
                    iALL(i) = iALL(UBound(iALL))
                    ReDim Preserve iALL(UBound(iALL) - 1)
                    Exit For
                End If
            Next
        Next
        Hmany = Hmany - TreeView1.SelectedItem.Children
        'LID = CStr(Hex(CLng("&H" & LID) - TreeView1.SelectedItem.Children))
        'LID = StrTo4Str(LID)
        nownum = Get_Index(Fk)
    End If
    '普通书签
    frmMain.StatusBar1.Panels.Item(2).Text = nownum
    TreeView1.Nodes.Remove (TreeView1.Nodes.Item(iALL(nownum).MyID & "A").Index)
    iALL(nownum) = iALL(UBound(iALL))
    ReDim Preserve iALL(UBound(iALL) - 1)
    Hmany = Hmany - 1
    'LID = CStr(Hex(CLng("&H" & LID) - 1))
    'LID = StrTo4Str(LID)
    
    nownum = Get_Index(TreeView1.Nodes.Item(1).Root.Key)
    Call Fash
End Sub
 '添加
Private Sub mFront_Click() '前面
    mAddSub (3)
End Sub
Private Sub mBehind_Click() '后面
    mAddSub (2)
End Sub
Private Sub mSub_Click() '子集
    mAddSub (4)
End Sub

Private Sub mAddSub(ty As Byte)  '添加(类型)
    Hmany = Hmany + 1
    LID = CStr(Hex(CLng("&H" & LID) + 1))
    LID = StrTo4Str(LID)
    ReDim Preserve iALL(Hmany)
    
    iALL(Hmany).CName = "新书签"
    iALL(Hmany).MyID = LID
    iALL(Hmany).Tyte = 4
    iALL(Hmany).WWW = "http://www.google.com.hk"
    
    If ty = 4 Then '添加到目标子级
        iALL(Hmany).FatherID = Left(TreeView1.SelectedItem.Key, 4)
    ElseIf ty = 1 Then
        iALL(Hmany).FatherID = "FFFF"
        TreeView1.Nodes.Add , , LID & "A", iALL(Hmany).CName
    ElseIf Not (TreeView1.SelectedItem.Parent Is Nothing) Then '添加到目标同级
        iALL(Hmany).FatherID = Left(TreeView1.SelectedItem.Parent.Key, 4)
    ElseIf TreeView1.SelectedItem.Parent Is Nothing Then '添加到目标同级
        iALL(Hmany).FatherID = "FFFF"
    Else
        MsgBox "特殊错误！请联系开发者！错误代号mAddSub", vbOKOnly, "警告": End
    End If
    If ty <> 1 Then TreeView1.Nodes.Add TreeView1.SelectedItem, ty, LID & "A", iALL(Hmany).CName
    nownum = Hmany
    Call Fash
End Sub




'×××××××××××××预览数据，拖动××××××××××××××
'2-准备右键菜单,1-选中
Private Sub TreeView1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then
    iButton(0) = 2: iButton(1) = X: iButton(2) = Y '标示为右键
ElseIf Button = 1 Then
    iButton(0) = 1 '标示为左键
End If
'显示元素内容
If Not (TreeView1.HitTest(X, Y) Is Nothing) Then
    nownum = Get_Index(TreeView1.HitTest(X, Y).Key) '取得数组编号
    frmMain.StatusBar1.Panels.Item(2).Text = iALL(nownum).CName
    
    'frmMain.TreeView1.SelectedItem = TreeView1.DropHighlight
    TreeView1.DropHighlight = TreeView1.HitTest(X, Y)
    TreeView1.DropHighlight.EnsureVisible '展开显示项
    'frmMain.StatusBar1.Panels.Item(2).Text = frmMain.StatusBar1.Panels.Item(2).Text & "MouseDown" & frmMain.TreeView1.SelectedItem

    '调整菜单显示
        '复位
    mTXT.Visible = True: mHTML.Visible = True: mUrl.Visible = True: mFavorites.Visible = True
    mOut.Enabled = True
    'mTXT.Enabled = False: mHTML.Enabled = False: mUrl.Enabled = False: mFavorites.Enabled = False

    frmMain.TreeView1.SelectedItem = frmMain.TreeView1.Nodes(iALL(nownum).MyID & "A")
    
    CmdAdd.Enabled = IIf(frmMain.TreeView1.SelectedItem Is Nothing, False, True)
    CmdDele.Enabled = IIf(frmMain.TreeView1.SelectedItem Is Nothing, False, True)
    CmdU.Enabled = IIf(frmMain.TreeView1.SelectedItem.Previous Is Nothing, False, True)
    CmdD.Enabled = IIf(frmMain.TreeView1.SelectedItem.Next Is Nothing, False, True)
    mUp.Enabled = CmdU.Enabled: mDown.Enabled = CmdD.Enabled

    mSub.Enabled = IIf(iALL(nownum).Tyte = 4, False, True) '非目录不能添加子集
    If frmMain.TreeView1.SelectedItem.Children = 0 And iALL(nownum).Tyte = 5 Then mOut.Enabled = False '目录无子集不能导出
    '预览数据
    txtName.Text = iALL(nownum).CName

    If TreeView1.Nodes.Item(iALL(nownum).MyID & "A").Children = 0 Then   '无子集
        txtWWW.Text = iALL(nownum).WWW
        txtWWW.Enabled = True
        CmdGO.Enabled = True
    Else
        txtWWW.Text = ""
        txtWWW.Enabled = False
        CmdGO.Enabled = False
    End If
    ChkType.Value = IIf(iALL(nownum).Tyte = 5, 1, 0)
End If
End Sub

'Private Sub TreeView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'    iButton(0) = 0
'    frmMain.TreeView1.SelectedItem = frmMain.TreeView1.Nodes(iALL(nownum).MyID & "A")
'End Sub

'选中元素&右键菜单
Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
'TreeView1.SelectedItem = TreeView1.Nodes.Item(iALL(nownum).MyID & "A")
'frmMain.TreeView1.SelectedItem = TreeView1.DropHighlight

If iButton(0) = 1 Then
    iButton(0) = 0
    '预览
    'If Not (TreeView1.HitTest(iButton(1), iButton(2)) Is Nothing) Then'nownum = Get_Index(TreeView1.HitTest(iButton(1), iButton(2)).Key)
ElseIf iButton(0) = 2 Then '右键菜单
    '右键无法使用全部导出功能
    mTXT.Visible = False: mHTML.Visible = False: mUrl.Visible = False: mFavorites.Visible = False
    frmMain.PopupMenu mEdit, vbPopupMenuLeftAlign, iButton(1) + TreeView1.Left, iButton(2) + TreeView1.Top
    iButton(0) = 1
End If
End Sub
'鼠标移动-拖动
Private Sub TreeView1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If frmMain.TreeView1.Nodes.Count <> 0 Then frmMain.TreeView1.SelectedItem = frmMain.TreeView1.Nodes(iALL(nownum).MyID & "A")
    If Button = vbLeftButton And Not (TreeView1.HitTest(X, Y) Is Nothing) Then '指示一个拖动操作。
        TreeView1.SelectedItem = TreeView1.Nodes(iALL(nownum).MyID & "A")
        '使用CreateDragImage方法设置拖动图标。
        TreeView1.DragIcon = TreeView1.SelectedItem.CreateDragImage
        TreeView1.Drag vbBeginDrag '拖动操作。
    Else
        'TreeView1.MousePointer = vbNoDrop
    End If
End Sub
'拖动停止
Private Sub TreeView1_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    If iButton(0) = 0 Then Exit Sub
    TreeView1.DropHighlight = TreeView1.HitTest(X, Y) '高亮
    If Not (TreeView1.DropHighlight Is Nothing) Then '打开折叠的
        If TreeView1.DropHighlight <> TreeView1.SelectedItem Then TreeView1.DropHighlight.Expanded = True
        TreeView1.DropHighlight.EnsureVisible
    End If
End Sub
'移动到高亮项子集
Private Sub TreeView1_DragDrop(Source As Control, X As Single, Y As Single)
If TreeView1.DropHighlight Is Nothing Then Exit Sub

'目标不是空；目标和选中不是一个
If TreeView1.DropHighlight Is Nothing Then Exit Sub
If Not (TreeView1.DropHighlight Is Nothing) And TreeView1.DropHighlight <> TreeView1.SelectedItem Then
    'Dim Highlight As String
    'Highlight = TreeView1.DropHighlight.Key
    'TreeView1.Nodes.Remove (TreeView1.SelectedItem.Index)
    'TreeView1.Nodes.Add Highlight, 4, element(nownum).name, element(nownum).name, element(nownum).Image
    
    
    '目标有上级（目标的上级是目录）
    If Not (TreeView1.DropHighlight.Parent Is Nothing) Then 'TreeView1.SelectedItem.Children <> 0 And
        MsgBox "你现在将一个书签(或目录)放到有上级的目录下，造成目录套嵌，UC浏览器目前无法识别这种目录套嵌格式。终止操作。", vbOKOnly, "无法操作"
        Exit Sub
    '选中目标是目录
    ElseIf TreeView1.SelectedItem.Children <> 0 Then
        MsgBox "你现在将一个目录放到另一个目录下，造成目录套嵌，UC浏览器目前无法识别这种目录套嵌格式。终止操作。", vbOKOnly, "无法操作"
        Exit Sub
    End If
    '将选中拖动到一个书签目标中
    If iALL(Get_Index(TreeView1.DropHighlight.Key)).Tyte <> 5 Then
        If MsgBox("你现在将一个书签更改为一个目录，次书签的网址部分将丢失？确认，改变类型；取消，终止操作。", vbOKCancel, "改变类型") = vbOK Then
            iALL(Get_Index(TreeView1.DropHighlight.Key)).WWW = ""
            iALL(Get_Index(TreeView1.DropHighlight.Key)).Tyte = 5
        Else
            Exit Sub
        End If
    End If
    iALL(nownum).FatherID = iALL(Get_Index(TreeView1.DropHighlight.Key)).MyID
    On Error GoTo checkerror
    Set TreeView1.SelectedItem.Parent = TreeView1.DropHighlight
End If

Set TreeView1.DropHighlight = Nothing
Exit Sub
checkerror:
    ' Define constants to represent Visual Basic errors code.
    Const CircularError = 35614
    If Err.Number = CircularError Then
        Dim msg As String
        msg = "A node can't be made a child of its own children."
    End If
End Sub







'×××××××××××××及时改变××××××××××××××
Private Sub TreeView1_AfterLabelEdit(Cancel As Integer, NewString As String)
    iALL(nownum).CName = NewString
    txtName.Text = iALL(nownum).CName
End Sub
Private Sub txtName_Change()
    iALL(nownum).CName = txtName.Text
    TreeView1.Nodes.Item(iALL(nownum).MyID & "A").Text = iALL(nownum).CName
    txtName.ToolTipText = txtName.Text
End Sub

Private Sub ChkType_Click()
    iALL(nownum).Tyte = IIf(ChkType.Value = 0, 4, 5)
    ChkType.ToolTipText = IIf(ChkType.Value = 0, "书签", "目录")
    If iALL(nownum).Tyte = 4 And TreeView1.Nodes.Item(iALL(nownum).MyID & "A").Children <> 0 Then
        '将目录设置为书签
        If MsgBox("你现在将一个存在子集的目录更改为一个书签？确认，删除子集；取消，终止操作。", vbOKCancel, "改变类型") = vbOK Then
            txtWWW.Text = "http://"
            txtWWW.Enabled = True
            Dim i As Integer
            For i = 1 To TreeView1.SelectedItem.Children
                TreeView1.Nodes.Remove (TreeView1.SelectedItem.Child.Index)
            Next
        Else
            ChkType.Value = 1
            Exit Sub
        End If
    ElseIf iALL(nownum).Tyte = 4 And TreeView1.Nodes.Item(iALL(nownum).MyID & "A").Children = 0 Then
        '将空目录设置为书签
        If txtWWW.Text = "" Then txtWWW.Text = "http://"
    ElseIf iALL(nownum).Tyte = 5 And iALL(nownum).WWW <> "" Then
        '将书签设置为目录
        If MsgBox("你现在将一个书签更改为一个目录，次书签的网址部分将丢失？确认，改变类型；取消，终止操作。", vbOKCancel, "改变类型") = vbOK Then
            txtWWW.Text = ""
            txtWWW.Enabled = False
        Else
            ChkType.Value = 0
            Exit Sub
        End If
    End If
End Sub

Private Sub txtWWW_Change()
    iALL(nownum).WWW = txtWWW.Text
    txtWWW.ToolTipText = txtWWW.Text
End Sub
