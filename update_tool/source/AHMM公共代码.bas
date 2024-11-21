Attribute VB_Name = "PublicSource"
Option Explicit

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, _
ByVal lpoperation As String, ByVal lpfile As String, ByVal lpparameters As String, _
ByVal lpdirectory As String, ByVal nshowcmd As Long) As Long

Public Declare Function timeGetTime Lib "winmm.dll" () As Long

'Public Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (LpBrowseInfo As BROWSEINFO) As Long
'Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long


Public Function DirectoryExists(ByVal dirPath As String) As Boolean
    ' ���dirPath�Ƿ��Է�б�ܽ�β��������ǣ������������ѡ�����Ƽ���
    If Right(dirPath, 1) <> "\" Then dirPath = dirPath & "\"
    
    ' ���Ի�ȡĿ¼�е��κ���Ŀ��ʹ��ͨ�����
    Dim anyEntry As String
    anyEntry = Dir(dirPath & "*") ' ��������ʹ�� "*" ������ vbDirectory����Ϊ����ֻ����֪���Ƿ����κ���Ŀ����
    
    ' ���Dir()�����˷ǿ��ַ�������Ŀ¼����
    DirectoryExists = (anyEntry <> "")
End Function
