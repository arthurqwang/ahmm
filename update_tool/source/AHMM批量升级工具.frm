VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form AHMMUpdateTool 
   Caption         =   "AHMM������������"
   ClientHeight    =   6660
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15885
   Icon            =   "AHMM������������.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6660
   ScaleWidth      =   15885
   StartUpPosition =   3  '����ȱʡ
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
      TextRTF         =   $"AHMM������������.frx":424A
   End
   Begin VB.CommandButton BT_LINK_WEB 
      Caption         =   $"AHMM������������.frx":42F1
      Height          =   495
      Left            =   2760
      TabIndex        =   13
      Top             =   5880
      Width           =   4215
   End
   Begin VB.CommandButton BT_END 
      Caption         =   "�˳�"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "ѡ��ģ��"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "���"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "��ʼ����"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "���ɵİ汾�����޷���������ʹ����ͼ�ļ��Դ����������ܵ�������"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "��ע�⡿"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   $"AHMM������������.frx":4322
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "����"
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
      Picture         =   "AHMM������������.frx":433F
      Stretch         =   -1  'True
      ToolTipText     =   "������ʡ���ϵͳ�ۡ���վ"
      Top             =   4320
      Width           =   1800
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   $"AHMM������������.frx":156C7
      Height          =   1455
      Left            =   2760
      TabIndex        =   12
      Top             =   4320
      Width           =   8655
   End
   Begin VB.Label Label6 
      Caption         =   "���ļ���������AHMM��������������������ļ��洢���������ļ���NewAHMM�С�"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "��ָ�����ļ������ļ��С�"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "ָ���°汾��AHMM�ļ�����Ϊ�ɰ汾������Ŀ�ꡣ������ahmm.html��Ҳ��ʹ�������κ�AHMM�ļ���"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "��ѡ��ģ���ļ���"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "��һ���°汾AHMM�ļ�Ϊģ�壬��һ���ļ���������AHMM�ļ�һ������"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "��ɫȫϢ��ͼAHMM������������"
      BeginProperty Font 
         Name            =   "����"
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
      Picture         =   "AHMM������������.frx":15843
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
'******************************************    AHMM ������������   *********************************************
'******************************************        Ver 1.0.2       *********************************************
'******************************************      ���ߣ���Ȩ        *********************************************
'******************************************      2024.11.15        *********************************************
'***************************************************************************************************************
'***************************************************************************************************************

'������Ĺ��ܣ���������AHMM����ɫȫϢ��ͼ���ļ�
'�������̣���ģ���ļ����������滻Ϊ��AHMM�����ݣ�Ȼ����滻���AHMMȫ���ı�д���µ�AHMM***.html������������ͼ�ļ���
'���ɵİ汾�����޷���������Ϊ�������ṹ���죬��ʱӦ��ʹ����ͼ�ļ��Դ����������ܵ�������




Option Explicit

Dim g_FileNameMod As String    'ģ���ļ�������
Dim g_FileNameOld As String    '�ɰ��ļ�������(��������)
Dim g_FileNameNew As String    '��������°��ļ������ƣ�����ļ�����ͬ�����ļ��洢���µ��ļ�����

Dim g_FileAllTextMod As String 'ģ���ļ���ȫ���ı�
Dim g_FileAllTextOld As String '���ļ���ȫ���ı�
Dim g_FileAllTextNew As String '���ļ���ȫ���ı�����д�����ļ�

Dim g_FolderPathMod As String  'ģ��AHMM�����ļ���
Dim g_FolderPathOld As String  '�ɰ�AHMM�����ļ���
Dim g_FolderPathNew As String  '�µ�AHMM�ļ��洢���ļ��У����ھɰ��ļ����н��� NewAHMM

Dim g_StartDataMark As String  'AHMM�ļ����������Ŀ�ʼ���
Dim g_EndDataMark As String    'AHMM�ļ����������Ľ������

Dim g_ModData As String  'ģ���ļ���������ȫ���ı�������ʵ�ʵ�Ҫ�����ľ��ļ��������滻
Dim g_OldData As String  '���ļ���������ȫ���ı�����д�����ļ���

Dim g_FinishNum As Integer   '������������ļ�����

Dim g_Fso As Object   '�ļ�������ΪVB6����ֱ�Ӷ�д utf-8 �ļ�����Ҫ�øö�����


'**************************************  ����ѡ���ļ��� ***************************************
'����
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

' ����
Private Const BIF_RETURNONLYFSDIRS As Long = &H1
Private Const MAX_PATH As Long = 260

'����������ѡ����ļ���
Private Function SelectFolder() As String
    Dim bi As BROWSEINFO
    Dim pidl As Long
    Dim folderPath As String * MAX_PATH
    Dim result As Long

    ' ��ʼ��BROWSEINFO�ṹ
    bi.hwndOwner = Me.hWnd ' ��ȡ��ǰ���ľ��
    bi.pidlRoot = 0& ' �����濪ʼ���
    bi.pszDisplayName = 0&
    bi.lpszTitle = "��ѡ��ɰ� AHMM ���ڵ��ļ���" ' �Ի������
    bi.ulFlags = BIF_RETURNONLYFSDIRS ' ֻ�����ļ�ϵͳ�ļ���
    bi.lpfnCallback = 0 ' ����Ҫ�ص�����
    bi.lParam = 0 ' ����Ҫ�������
    bi.iImage = 0 ' ����Ҫͼ������

    ' ��ʾ�Ի���
    pidl = SHBrowseForFolder(bi)

    ' ����Ƿ�ѡ�����ļ���
    If pidl <> 0 Then
        ' ��ȡ�ļ���·��
        result = SHGetPathFromIDList(pidl, folderPath)
        If result Then
            SelectFolder = Left$(folderPath, InStr(folderPath, vbNullChar) - 1) ' ȥ��ĩβ�Ŀ��ַ�
        Else
            SelectFolder = ""
        End If
    Else
        SelectFolder = ""
    End If
End Function


'*************************************���� UTF-8 *************************************************
'VB6 ��д�ļ����� ANSI����UTF-8 ��Ҫ���⴦��ahmm***.html ��UTF-8 ��ʽ��
'����Ҫ����  Microsoft ActiveX Data Objects 2.8����������ͨ�÷����������ģ����
'���������ɵ��ı��ļ�����BOM����ԭ����ahmm***.html������BOM��û��Ӱ�죬���ش����Ҵ����Ѷȴ�

'����ΪUTF8��ʽ���ı�
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

'��UTF-8��ʽ����ȫ���ı�
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


'��ʼ��
Private Sub Form_Load()
    
    '����뱾����ͬ�ļ����´���ahmm.html����ģ���ļ�Ĭ����Ϊ��
    If Dir(App.Path & "\ahmm.html") <> "" Then
        TXT_MOD_FILE.text = App.Path & "\ahmm.html"
    Else
        TXT_MOD_FILE.text = ""
    End If
    
    '���ļ���Ĭ����Ϊ��ģ����ͬ�����水������
    TXT_OLD_FOLDER.text = App.Path
    
    '�����ļ��Ի���ؼ��ļ� COMDLG32.OCX��ָ��ģ���ļ��ĶԻ�����Ҫ����
    '��Ϊ����ϵͳ�п���û������ļ�������Ҫ���һ�£����û���򿽱���ȥ
    '����ļ�Ҫ�����������߷���ͬһ��ѹ�������ļ�����
    CopyCOMDLG32OCX

End Sub



'ָ��ģ���ļ�
'�ļ��Ի���ؼ��ļ� COMDLG32.OCX
'��Ϊ����ϵͳ�п���û������ļ�������Ҫ���һ�£����û���򿽱���ȥ��ǰ���Ѵ���
'����ļ�Ҫ�����������߷���ͬһ��ѹ�������ļ�����
Private Sub BT_SELECT_MOD_FILE_Click()
    
    Dim lastBackslashPos As String

    CommonDialog1.Filter = "HTML Files (*.htm; *.html)|*.htm;*.html|All Files (*.*)|*.*"
    Me.CommonDialog1.ShowOpen
    If CommonDialog1.FileName <> "" Then
        TXT_MOD_FILE.text = CommonDialog1.FileName
    End If
    
    '�Զ����þ��ļ���·����ģ���ļ���ͬ��ȥ�����һ��\���Ժ���ַ�
    '�������\
    lastBackslashPos = InStrRev(TXT_MOD_FILE.text, "\")
  
    ' ����ҵ��˷�б�ܣ����ȡ�ַ�������λ��֮ǰ
    If lastBackslashPos > 0 Then
        TXT_OLD_FOLDER.text = Left(TXT_MOD_FILE.text, lastBackslashPos - 1)
    End If

End Sub


'ָ�����ļ������ļ���
Private Sub BT_BROWSE_OLD_FOLDER_Click()
    Dim t As String
    t = SelectFolder()
    If t <> "" Then TXT_OLD_FOLDER.text = t
End Sub



'�����������ļ�
Private Sub BT_GO_Click()

    'AHMM�ļ�����������ֹ��־
    g_StartDataMark = "<div id=""DATA_SAVER"" class=""data_saver"">"
    g_EndDataMark = "</div>   <!-- <div id=""DATA_SAVER"">���� -->"
    
    '�����ȡ����ģ���ļ�������������
    g_ModData = ""
    
    
    '********************* ����ģ���ļ������ȫ���ı���������������ݣ����滻��*******************************
    '*********************************************************************************************************
    
    '���ģ���ļ�·��
    g_FileNameMod = Trim(TXT_MOD_FILE.text)    '����ȫ·�����ļ���
    
    g_FolderPathMod = g_FileNameMod
    'ȥ��β��\���������
    g_FolderPathMod = CutLastBackslash(g_FolderPathMod)
    
    '���δָ��ģ���ļ�������ʾ�����ж�
    If g_FileNameMod = "" Then
        MsgBox "��ģ���ļ�δָ��������ָ��һ���°汾�İ�ɫȫϢ��ͼ AHMM �ļ�����Ϊ����ģ�塣" & vbCrLf & "������ָ����", vbCritical
        Exit Sub
    End If
    
    '���ģ���ļ������ڣ�����ʾ�����ж�
    If Dir(g_FileNameMod) = "" Then
        MsgBox "���ļ�δ�ҵ�����ָ���İ�ɫȫϢ��ͼ AHMM ģ���ļ�δ�ҵ���" & vbCrLf & "������ָ����", vbCritical
        Exit Sub
    End If
    
    '��ȡģ���ļ�ȫ���ı�����Ϊ��utf-8��ʽ�����Բ�������ͨ��open for input����ͨ��ֻ֧��ANSI��д��Ҳһ����
    g_FileAllTextMod = LoadAsUTF8(g_FileNameMod)
    '��ȡģ���ļ��е����� g_ModData���������滻Ϊ�������ļ�������
    g_ModData = ExtractStringBetween(g_FileAllTextMod, g_StartDataMark, g_EndDataMark)
    
    '���ģ���ļ��в�����������������ʾ�����ж�
    If Trim(g_ModData) = "" Then
        MsgBox "����ʽ���󡿣�ָ���İ�ɫȫϢ��ͼ AHMM ģ���ļ���ʽ����ȷ��" & vbCrLf & " ������ָ����ȷ��ʽ�� AHMM �ļ���", vbCritical
        Exit Sub
    End If
    
    
    '********************* ���δ���������ļ������ȫ���ı���������������ݣ�ȥ�滻ģ�壩*********************
    '*********************************************************************************************************
    
    '���þ��ļ�·��
    g_FolderPathOld = Trim(TXT_OLD_FOLDER.text)
    'ȥ��β��\���������
    g_FolderPathOld = CutLastBackslash(g_FolderPathOld)
    
    '���δָ���ɰ��ļ��У�����ʾ�����ж�
    If g_FolderPathOld = "" Then
        MsgBox "���ɰ��ļ���δָ��������ָ���ɰ汾��ͼ�����ļ��У����ļ����е����� AHMM �ļ�������������" & vbCrLf & "������ָ����", vbCritical
        Exit Sub
    End If
    
    '����ɰ��ļ��в����ڣ�����ʾ�����ж�
    If DirectoryExists(g_FolderPathOld) = False Then
        MsgBox "���ɰ��ļ���δ�ҵ���������ȷָ���ɰ汾��ͼ�����ļ��У����ļ����е����� AHMM �ļ�������������" & vbCrLf & "������ָ����", vbCritical
        Exit Sub
    End If
    
    '�����������AHMM���ļ��� NewAHMM
    Set g_Fso = CreateObject("Scripting.FileSystemObject")
    g_FolderPathNew = g_FolderPathOld & "\NewAHMM"
    'ȥ��β��\���������
    g_FolderPathNew = CutLastBackslash(g_FolderPathNew)
    
    '���NewAHMM�ļ��в����ڣ��򴴽���
    If Not g_Fso.FolderExists(g_FolderPathNew) Then
        g_Fso.CreateFolder g_FolderPathNew
    End If
    
    
    '��õ�һ���ļ���
    g_FileNameOld = Dir(g_FolderPathOld & "\*.html")
    
    '�����ɹ�������
    g_FinishNum = 0
    
    '��ʾ����ʼ
    TXT_PROGRESS.text = TXT_PROGRESS.text & vbCrLf & vbCrLf & "========================================" & vbCrLf & "������ʼ��" & vbCrLf
    
    'ѭ������������ļ�
    Do While g_FileNameOld <> ""
    
        If Trim(LCase(g_FolderPathOld & "\" & g_FileNameOld)) <> Trim(LCase(g_FileNameMod)) Then   '�����ģ���ļ�����Խ��
        
            '��ȡ���ļ�ȫ���ı���utf-8��ʽ
            g_FileAllTextOld = LoadAsUTF8(g_FolderPathOld & "\" & g_FileNameOld)
            
            '��ȡ�ɰ汾�ļ��е����� g_OldData
            g_OldData = ExtractStringBetween(g_FileAllTextOld, g_StartDataMark, g_EndDataMark)
            
            '���g_OldDataΪ�գ����ʾ���ⲻ��һ��AHMM�ļ���������
            '��Ϊ�գ����ʾ��AHMM
            If g_OldData <> "" Then
            
                '��ģ���е������滻Ϊ��������AHMM������
                g_FileAllTextNew = Replace(g_FileAllTextMod, g_ModData, g_OldData)
                
                '�����°�� AHMM �ļ���uft-8 ��ʽ�������� NewAHMM �ļ����У��ļ�������
                g_FileNameNew = g_FileNameOld
                Call SaveAsUTF8(g_FileAllTextNew, g_FolderPathNew & "\" & g_FileNameNew)
                
                '�����ɹ�����1
                g_FinishNum = g_FinishNum + 1
                
                '��ʾ�ɹ����ļ���
                TXT_PROGRESS.text = TXT_PROGRESS.text & vbCrLf & " " & g_FinishNum & ": " & g_FileNameNew & ": �����ɹ�"
                '�������Ϊ����ʾ����
                TXT_PROGRESS.SelStart = Len(TXT_PROGRESS.text) ' ����ѡ��ʼλ��Ϊ�ı���ĩβ
                TXT_PROGRESS.SelLength = 0 ' ����ѡ�񳤶�Ϊ0�������λ��
                TXT_PROGRESS.Refresh ' ˢ���ı�����ʾ
                
                MySleep 5
                
            End If
                
        End If
        
        g_FileNameOld = Dir()   '��ȡ��һ�����ļ�
        
    Loop
    
    '��ʾ�������
    TXT_PROGRESS.text = TXT_PROGRESS.text & vbCrLf & vbCrLf & "������������ɹ����� " & g_FinishNum & " ����" & vbCrLf & vbCrLf
    TXT_PROGRESS.text = TXT_PROGRESS.text & "��������ļ��洢��NewAHMM���ļ��У�����" & vbCrLf & g_FolderPathNew & vbCrLf & vbCrLf & "����ʾ�����������������ʹ�ñ������������ɵİ汾���������⣬��ʹ����ͼ�ļ��Դ����������ܵ�������" & vbCrLf
    '�������Ϊ����ʾ����
    TXT_PROGRESS.SelStart = Len(TXT_PROGRESS.text) ' ����ѡ��ʼλ��Ϊ�ı���ĩβ
    TXT_PROGRESS.SelLength = 0 ' ����ѡ�񳤶�Ϊ0�������λ��
    TXT_PROGRESS.Refresh ' ˢ���ı�����ʾ
 
End Sub



'�������˳�
Private Sub BT_END_Click()
    End
End Sub



'���ӵ���ϵͳ����վ
Private Sub BT_LINK_WEB_Click()
    Dim lngReturn As Long
    lngReturn = ShellExecute(Me.hWnd, "open", "http://www.holomind.com.cn", "", "", 0)
End Sub



' ��ȡ�����ַ���֮�������
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
  


'·���ַ�����������һ���ַ���\����ȥ������Ϊϵͳ��ʱ�����\��ͳһ��������������\
Function CutLastBackslash(ByVal s As String) As String
    If Right(s, 1) = "\" Then
        ' ����ǣ���ȥ����
        CutLastBackslash = Left(s, Len(s) - 1)
    Else
        ' ������ǣ��򱣳�ԭ��
        CutLastBackslash = s
    End If
End Function



'��ͣ MySsleep
Public Sub MySleep(ms As Long) 'ԭSleep����������Ȩ������һ���� ms:������
    Dim BeginTime As Long
    BeginTime = timeGetTime '���¿�ʼʱ��ʱ��
    While timeGetTime < BeginTime + ms 'ѭ���ȴ�
        DoEvents 'ת�ÿ���Ȩ���Ա��ò���ϵͳ�����������¼�
    Wend
End Sub



'�����ļ��Ի���ؼ��ļ� COMDLG32.OCX����Ϊ����ϵͳ�п���û������ļ�������Ҫ���һ�£����û���򿽱���ȥ
'����ļ�Ҫ�����������߷���ͬһ��ѹ�������ļ�����
Sub CopyCOMDLG32OCX()
    Dim tempStr, strFileName As String

    '����COMDLG32.OCX�� C:\Windows\System32\ ������ļ��Ĺ�����Ҫ���ļ�
    tempStr = IIf(Len(App.Path) > 3, App.Path & "\COMDLG32.OCX", App.Path & "COMDLG32.OCX")
    strFileName = "C:\Windows\System32\COMDLG32.OCX"
      
    If Dir(tempStr) <> "" And Dir(strFileName) = "" Then
        FileCopy tempStr, strFileName
    End If
End Sub



























