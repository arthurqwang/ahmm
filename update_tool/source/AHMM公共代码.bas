Attribute VB_Name = "PublicSource"
Option Explicit

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, _
ByVal lpoperation As String, ByVal lpfile As String, ByVal lpparameters As String, _
ByVal lpdirectory As String, ByVal nshowcmd As Long) As Long

Public Declare Function timeGetTime Lib "winmm.dll" () As Long

'Public Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (LpBrowseInfo As BROWSEINFO) As Long
'Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long


Public Function DirectoryExists(ByVal dirPath As String) As Boolean
    ' 检查dirPath是否以反斜杠结尾，如果不是，则添加它（可选，但推荐）
    If Right(dirPath, 1) <> "\" Then dirPath = dirPath & "\"
    
    ' 尝试获取目录中的任何条目（使用通配符）
    Dim anyEntry As String
    anyEntry = Dir(dirPath & "*") ' 这里我们使用 "*" 而不是 vbDirectory，因为我们只是想知道是否有任何条目返回
    
    ' 如果Dir()返回了非空字符串，则目录存在
    DirectoryExists = (anyEntry <> "")
End Function
