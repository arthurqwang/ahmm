VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form AHMMUpdateTool 
   Caption         =   "AHMM批量升级工具"
   ClientHeight    =   6660
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15885
   Icon            =   "AHMM批量升级工具.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6660
   ScaleWidth      =   15885
   StartUpPosition =   3  '窗口缺省
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox TXT_PROGRESS 
      Height          =   5535
      Left            =   11520
      TabIndex        =   15
      Top             =   840
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   9763
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"AHMM批量升级工具.frx":424A
   End
   Begin VB.CommandButton BT_LINK_WEB 
      Caption         =   $"AHMM批量升级工具.frx":42F1
      Height          =   495
      Left            =   2760
      TabIndex        =   13
      Top             =   5880
      Width           =   4215
   End
   Begin VB.CommandButton BT_END 
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   11
      Top             =   5880
      Width           =   4215
   End
   Begin VB.TextBox TXT_OLD_FOLDER 
      Height          =   375
      Left            =   680
      TabIndex        =   10
      Top             =   2540
      Width           =   8895
   End
   Begin VB.TextBox TXT_MOD_FILE 
      Height          =   375
      Left            =   680
      TabIndex        =   7
      Top             =   1700
      Width           =   8895
   End
   Begin VB.CommandButton BT_SELECT_MOD_FILE 
      Caption         =   "选择模板"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9600
      TabIndex        =   4
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton BT_BROWSE_OLD_FOLDER 
      Caption         =   "浏览"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9600
      TabIndex        =   3
      Top             =   2540
      Width           =   1695
   End
   Begin VB.CommandButton BT_GO 
      BackColor       =   &H00C0C0FF&
      Caption         =   "开始升级"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   680
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3360
      Width           =   10620
   End
   Begin VB.Label Label11 
      Caption         =   "过旧的版本可能无法升级，可使用脑图文件自带的升级功能单个处理。"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1500
      TabIndex        =   18
      Top             =   2950
      Width           =   8415
   End
   Begin VB.Label Label10 
      Caption         =   "【注意】"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   17
      Top             =   2950
      Width           =   975
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   $"AHMM批量升级工具.frx":4322
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9360
      TabIndex        =   16
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "升级进度"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12240
      TabIndex        =   14
      Top             =   460
      Width           =   2655
   End
   Begin VB.Image IMG_BSV 
      Appearance      =   0  'Flat
      Height          =   2055
      Left            =   680
      Picture         =   "AHMM批量升级工具.frx":433F
      Stretch         =   -1  'True
      ToolTipText     =   "点击访问【大系统观】网站"
      Top             =   4320
      Width           =   1800
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   $"AHMM批量升级工具.frx":156C7
      Height          =   1455
      Left            =   2760
      TabIndex        =   12
      Top             =   4320
      Width           =   8655
   End
   Begin VB.Label Label6 
      Caption         =   "此文件夹中所有AHMM都将被升级，升级后的文件存储在其下子文件夹NewAHMM中。"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   9
      Top             =   2280
      Width           =   8415
   End
   Begin VB.Label Label5 
      Caption         =   "【指定旧文件所在文件夹】"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   8
      Top             =   2280
      Width           =   2775
   End
   Begin VB.Label Label4 
      Caption         =   "指定新版本的AHMM文件，作为旧版本升级的目标。可以是ahmm.html，也可使用其他任何AHMM文件。"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   1440
      Width           =   9615
   End
   Begin VB.Label Label3 
      Caption         =   "【选择模板文件】"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "以一个新版本AHMM文件为模板，将一个文件夹中所有AHMM文件一键升级"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   960
      Width           =   7095
   End
   Begin VB.Label Label1 
      Caption         =   "阿色全息脑图AHMM批量升级工具"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   360
      Width           =   6615
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   480
      Picture         =   "AHMM批量升级工具.frx":15843
      Stretch         =   -1  'True
      Top             =   340
      Width           =   1980
   End
End
Attribute VB_Name = "AHMMUpdateTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



'***************************************************************************************************************
'***************************************************************************************************************
'******************************************    AHMM 批量升级工具   *********************************************
'******************************************        Ver 1.0.2       *********************************************
'******************************************      作者：王权        *********************************************
'******************************************      2024.11.15        *********************************************
'***************************************************************************************************************
'***************************************************************************************************************

'本程序的功能：批量升级AHMM（阿色全息脑图）文件
'工作过程：把模板文件的数据区替换为旧AHMM的数据，然后把替换后的AHMM全部文本写入新的AHMM***.html，重新生成脑图文件。
'过旧的版本可能无法升级，因为数据区结构差异，此时应该使用脑图文件自带的升级功能单个处理。




Option Explicit

Dim g_FileNameMod As String    '模板文件的名称
Dim g_FileNameOld As String    '旧板文件的名称(欲升级的)
Dim g_FileNameNew As String    '升级后的新板文件的名称，与旧文件名相同，新文件存储在新的文件夹中

Dim g_FileAllTextMod As String '模板文件的全部文本
Dim g_FileAllTextOld As String '旧文件的全部文本
Dim g_FileAllTextNew As String '新文件的全部文本，将写入新文件

Dim g_FolderPathMod As String  '模板AHMM所在文件夹
Dim g_FolderPathOld As String  '旧版AHMM所在文件夹
Dim g_FolderPathNew As String  '新的AHMM文件存储的文件夹，将在旧版文件夹中建立 NewAHMM

Dim g_StartDataMark As String  'AHMM文件中数据区的开始标记
Dim g_EndDataMark As String    'AHMM文件中数据区的结束标记

Dim g_ModData As String  '模板文件数据区的全部文本，将被实际的要升级的旧文件的数据替换
Dim g_OldData As String  '旧文件数据区的全部文本，将写入新文件中

Dim g_FinishNum As Integer   '完成了升级的文件总数

Dim g_Fso As Object   '文件对象。因为VB6不能直接读写 utf-8 文件，需要用该对象处理


'**************************************  用于选择文件夹 ***************************************
'声明
Private Type BROWSEINFO
    hwndOwner As Long
    pidlRoot As Long
    pszDisplayName As Long
    lpszTitle As String
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpbi As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long

' 常量
Private Const BIF_RETURNONLYFSDIRS As Long = &H1
Private Const MAX_PATH As Long = 260

'函数：返回选择的文件夹
Private Function SelectFolder() As String
    Dim bi As BROWSEINFO
    Dim pidl As Long
    Dim folderPath As String * MAX_PATH
    Dim result As Long

    ' 初始化BROWSEINFO结构
    bi.hwndOwner = Me.hWnd ' 获取当前表单的句柄
    bi.pidlRoot = 0& ' 从桌面开始浏览
    bi.pszDisplayName = 0&
    bi.lpszTitle = "请选择旧版 AHMM 所在的文件夹" ' 对话框标题
    bi.ulFlags = BIF_RETURNONLYFSDIRS ' 只返回文件系统文件夹
    bi.lpfnCallback = 0 ' 不需要回调函数
    bi.lParam = 0 ' 不需要额外参数
    bi.iImage = 0 ' 不需要图标索引

    ' 显示对话框
    pidl = SHBrowseForFolder(bi)

    ' 检查是否选择了文件夹
    If pidl <> 0 Then
        ' 获取文件夹路径
        result = SHGetPathFromIDList(pidl, folderPath)
        If result Then
            SelectFolder = Left$(folderPath, InStr(folderPath, vbNullChar) - 1) ' 去除末尾的空字符
        Else
            SelectFolder = ""
        End If
    Else
        SelectFolder = ""
    End If
End Function


'*************************************处理 UTF-8 *************************************************
'VB6 读写文件都是 ANSI，对UTF-8 需要特殊处理。ahmm***.html 是UTF-8 格式的
'工程要引用  Microsoft ActiveX Data Objects 2.8，下面两个通用方法建议放在模块中
'本程序生成的文本文件带有BOM，而原来的ahmm***.html不带有BOM。没有影响，不必处理，且处理难度大。

'保存为UTF8格式的文本
Sub SaveAsUTF8(ByVal text As String, ByVal FileName As String)
  Dim oStream As ADODB.Stream
  Set oStream = New ADODB.Stream
  oStream.Open
  oStream.Charset = "UTF-8"
  oStream.Type = adTypeText
  oStream.WriteText text
  oStream.SaveToFile FileName, adSaveCreateOverWrite
  oStream.Close
End Sub

'以UTF-8格式读入全部文本
Function LoadAsUTF8(ByVal FileName As String) As String
  Dim oStream As ADODB.Stream
  Set oStream = New ADODB.Stream
  oStream.Open
  oStream.Charset = "UTF-8"
  oStream.LoadFromFile FileName
  LoadAsUTF8 = oStream.ReadText()
  oStream.Close
End Function

'************************************************************************************************
'************************************************************************************************
'************************************************************************************************
'************************************************************************************************


'初始化
Private Sub Form_Load()
    
    '如果与本程序同文件夹下存在ahmm.html，则模板文件默认设为它
    If Dir(App.Path & "\ahmm.html") <> "" Then
        TXT_MOD_FILE.text = App.Path & "\ahmm.html"
    Else
        TXT_MOD_FILE.text = ""
    End If
    
    '旧文件夹默认设为与模板相同，后面按此设置
    TXT_OLD_FOLDER.text = App.Path
    
    '拷贝文件对话框控件文件 COMDLG32.OCX。指定模板文件的对话框需要它。
    '因为操作系统中可能没有这个文件，所以要检测一下，如果没有则拷贝过去
    '这个文件要跟本升级工具放在同一个压缩包或文件夹内
    CopyCOMDLG32OCX

End Sub



'指定模板文件
'文件对话框控件文件 COMDLG32.OCX
'因为操作系统中可能没有这个文件，所以要检测一下，如果没有则拷贝过去。前面已处理。
'这个文件要跟本升级工具放在同一个压缩包或文件夹内
Private Sub BT_SELECT_MOD_FILE_Click()
    
    Dim lastBackslashPos As String

    CommonDialog1.Filter = "HTML Files (*.htm; *.html)|*.htm;*.html|All Files (*.*)|*.*"
    Me.CommonDialog1.ShowOpen
    If CommonDialog1.FileName <> "" Then
        TXT_MOD_FILE.text = CommonDialog1.FileName
    End If
    
    '自动设置旧文件的路径与模板文件相同，去掉最后一个\及以后的字符
    '反向查找\
    lastBackslashPos = InStrRev(TXT_MOD_FILE.text, "\")
  
    ' 如果找到了反斜杠，则截取字符串到该位置之前
    If lastBackslashPos > 0 Then
        TXT_OLD_FOLDER.text = Left(TXT_MOD_FILE.text, lastBackslashPos - 1)
    End If

End Sub


'指定旧文件所在文件夹
Private Sub BT_BROWSE_OLD_FOLDER_Click()
    Dim t As String
    t = SelectFolder()
    If t <> "" Then TXT_OLD_FOLDER.text = t
End Sub



'升级各个旧文件
Private Sub BT_GO_Click()

    'AHMM文件中数据区起止标志
    g_StartDataMark = "<div id=""DATA_SAVER"" class=""data_saver"">"
    g_EndDataMark = "</div>   <!-- <div id=""DATA_SAVER"">结束 -->"
    
    '清空提取出的模板文件的数据区变量
    g_ModData = ""
    
    
    '********************* 处理模板文件，获得全部文本，获得数据区内容（待替换）*******************************
    '*********************************************************************************************************
    
    '获得模板文件路径
    g_FileNameMod = Trim(TXT_MOD_FILE.text)    '包含全路径和文件名
    
    g_FolderPathMod = g_FileNameMod
    '去掉尾部\，避免出错
    g_FolderPathMod = CutLastBackslash(g_FolderPathMod)
    
    '如果未指定模板文件，则提示，并中断
    If g_FileNameMod = "" Then
        MsgBox "【模板文件未指定】：请指定一个新版本的阿色全息脑图 AHMM 文件，作为升级模板。" & vbCrLf & "请重新指定。", vbCritical
        Exit Sub
    End If
    
    '如果模板文件不存在，则提示，并中断
    If Dir(g_FileNameMod) = "" Then
        MsgBox "【文件未找到】：指定的阿色全息脑图 AHMM 模板文件未找到。" & vbCrLf & "请重新指定。", vbCritical
        Exit Sub
    End If
    
    '读取模板文件全部文本。因为是utf-8格式，所以不能用普通的open for input，普通的只支持ANSI。写入也一样。
    g_FileAllTextMod = LoadAsUTF8(g_FileNameMod)
    '提取模板文件中的数据 g_ModData，它将被替换为各个旧文件的数据
    g_ModData = ExtractStringBetween(g_FileAllTextMod, g_StartDataMark, g_EndDataMark)
    
    '如果模板文件中不存在数据区，则提示，并中断
    If Trim(g_ModData) = "" Then
        MsgBox "【格式错误】：指定的阿色全息脑图 AHMM 模板文件格式不正确。" & vbCrLf & " 请重新指定正确格式的 AHMM 文件。", vbCritical
        Exit Sub
    End If
    
    
    '********************* 依次处理各个旧文件，获得全部文本，获得数据区内容（去替换模板）*********************
    '*********************************************************************************************************
    
    '设置旧文件路径
    g_FolderPathOld = Trim(TXT_OLD_FOLDER.text)
    '去掉尾部\，避免出错
    g_FolderPathOld = CutLastBackslash(g_FolderPathOld)
    
    '如果未指定旧版文件夹，则提示，并中断
    If g_FolderPathOld = "" Then
        MsgBox "【旧版文件夹未指定】：请指定旧版本脑图所在文件夹，该文件夹中的所有 AHMM 文件都将被升级。" & vbCrLf & "请重新指定。", vbCritical
        Exit Sub
    End If
    
    '如果旧版文件夹不存在，则提示，并中断
    If DirectoryExists(g_FolderPathOld) = False Then
        MsgBox "【旧版文件夹未找到】：请正确指定旧版本脑图所在文件夹，该文件夹中的所有 AHMM 文件都将被升级。" & vbCrLf & "请重新指定。", vbCritical
        Exit Sub
    End If
    
    '创建升级后的AHMM的文件夹 NewAHMM
    Set g_Fso = CreateObject("Scripting.FileSystemObject")
    g_FolderPathNew = g_FolderPathOld & "\NewAHMM"
    '去掉尾部\，避免出错
    g_FolderPathNew = CutLastBackslash(g_FolderPathNew)
    
    '如果NewAHMM文件夹不存在，则创建它
    If Not g_Fso.FolderExists(g_FolderPathNew) Then
        g_Fso.CreateFolder g_FolderPathNew
    End If
    
    
    '获得第一个文件名
    g_FileNameOld = Dir(g_FolderPathOld & "\*.html")
    
    '升级成功的总数
    g_FinishNum = 0
    
    '显示处理开始
    TXT_PROGRESS.text = TXT_PROGRESS.text & vbCrLf & vbCrLf & "========================================" & vbCrLf & "【处理开始】" & vbCrLf
    
    '循环处理各个旧文件
    Do While g_FileNameOld <> ""
    
        If Trim(LCase(g_FolderPathOld & "\" & g_FileNameOld)) <> Trim(LCase(g_FileNameMod)) Then   '如果是模板文件，则越过
        
            '读取旧文件全部文本，utf-8格式
            g_FileAllTextOld = LoadAsUTF8(g_FolderPathOld & "\" & g_FileNameOld)
            
            '提取旧版本文件中的数据 g_OldData
            g_OldData = ExtractStringBetween(g_FileAllTextOld, g_StartDataMark, g_EndDataMark)
            
            '如果g_OldData为空，则表示：这不是一个AHMM文件，则跳过
            '不为空，则表示是AHMM
            If g_OldData <> "" Then
            
                '将模板中的数据替换为被升级的AHMM的数据
                g_FileAllTextNew = Replace(g_FileAllTextMod, g_ModData, g_OldData)
                
                '生成新版的 AHMM 文件。uft-8 格式。保存在 NewAHMM 文件夹中，文件名不变
                g_FileNameNew = g_FileNameOld
                Call SaveAsUTF8(g_FileAllTextNew, g_FolderPathNew & "\" & g_FileNameNew)
                
                '升级成功数加1
                g_FinishNum = g_FinishNum + 1
                
                '显示成功的文件名
                TXT_PROGRESS.text = TXT_PROGRESS.text & vbCrLf & " " & g_FinishNum & ": " & g_FileNameNew & ": 升级成功"
                '下面代码为了显示到底
                TXT_PROGRESS.SelStart = Len(TXT_PROGRESS.text) ' 设置选择开始位置为文本的末尾
                TXT_PROGRESS.SelLength = 0 ' 设置选择长度为0，即光标位置
                TXT_PROGRESS.Refresh ' 刷新文本框显示
                
                MySleep 5
                
            End If
                
        End If
        
        g_FileNameOld = Dir()   '获取下一个旧文件
        
    Loop
    
    '显示处理结束
    TXT_PROGRESS.text = TXT_PROGRESS.text & vbCrLf & vbCrLf & "【处理结束：成功升级 " & g_FinishNum & " 个】" & vbCrLf & vbCrLf
    TXT_PROGRESS.text = TXT_PROGRESS.text & "升级后的文件存储在NewAHMM子文件夹，即：" & vbCrLf & g_FolderPathNew & vbCrLf & vbCrLf & "【提示】：请检查升级结果。使用本工具升级过旧的版本可能有问题，可使用脑图文件自带的升级功能单个处理。" & vbCrLf
    '下面代码为了显示到底
    TXT_PROGRESS.SelStart = Len(TXT_PROGRESS.text) ' 设置选择开始位置为文本的末尾
    TXT_PROGRESS.SelLength = 0 ' 设置选择长度为0，即光标位置
    TXT_PROGRESS.Refresh ' 刷新文本框显示
 
End Sub



'结束，退出
Private Sub BT_END_Click()
    End
End Sub



'连接到大系统观网站
Private Sub BT_LINK_WEB_Click()
    Dim lngReturn As Long
    lngReturn = ShellExecute(Me.hWnd, "open", "http://www.holomind.com.cn", "", "", 0)
End Sub



' 提取两个字符串之间的内容
Function ExtractStringBetween(text As String, startTag As String, endTag As String) As String
    Dim startPos As Long
    Dim endPos As Long
    startPos = InStr(text, startTag) + Len(startTag)
    If startPos > Len(startTag) Then
        endPos = InStr(startPos, text, endTag)
        If endPos > startPos Then
            ExtractStringBetween = Mid(text, startPos, endPos - startPos)
        End If
    End If
End Function
  


'路径字符串中如果最后一个字符是\，则去掉，因为系统有时会带有\，统一处理避免出现两个\
Function CutLastBackslash(ByVal s As String) As String
    If Right(s, 1) = "\" Then
        ' 如果是，则去掉它
        CutLastBackslash = Left(s, Len(s) - 1)
    Else
        ' 如果不是，则保持原样
        CutLastBackslash = s
    End If
End Function



'暂停 MySsleep
Public Sub MySleep(ms As Long) '原Sleep不交出控制权，改造一个。 ms:毫秒数
    Dim BeginTime As Long
    BeginTime = timeGetTime '记下开始时的时间
    While timeGetTime < BeginTime + ms '循环等待
        DoEvents '转让控制权，以便让操作系统处理其它的事件
    Wend
End Sub



'拷贝文件对话框控件文件 COMDLG32.OCX。因为操作系统中可能没有这个文件，所以要检测一下，如果没有则拷贝过去
'这个文件要跟本升级工具放在同一个压缩包或文件夹内
Sub CopyCOMDLG32OCX()
    Dim tempStr, strFileName As String

    '拷贝COMDLG32.OCX到 C:\Windows\System32\ 。浏览文件的功能需要该文件
    tempStr = IIf(Len(App.Path) > 3, App.Path & "\COMDLG32.OCX", App.Path & "COMDLG32.OCX")
    strFileName = "C:\Windows\System32\COMDLG32.OCX"
      
    If Dir(tempStr) <> "" And Dir(strFileName) = "" Then
        FileCopy tempStr, strFileName
    End If
End Sub



























