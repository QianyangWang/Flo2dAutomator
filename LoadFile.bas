Attribute VB_Name = "LoadFile"
Public Const BIF_RETURNONLYFSDIRS = 1
Public Const BIF_DONTGOBELOWDOMAIN = 2
Public Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (LpBrowseInfo As BROWSEINFO) As Long
Public Declare Function SHGetPathFromIDlist Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Public Type BROWSEINFO
hOwner As Long
pidlroot As Long
pszDisplayName As String
lpszTitle As String
ulFlags As Long
lpfn As Long
lparam As Long
iImage As Long
End Type
Public Function GetFolder(ByVal hWnd As Long, Optional Title As String) As String
    Dim bi As BROWSEINFO
    Dim pidl As Long
    Dim folder As String
    folder = Space(255)
With bi
   If IsNumeric(hWnd) Then .hOwner = hWnd
   .ulFlags = BIF_RETURNONLYFSDIRS
   .pidlroot = 0
   If Title <> "" Then
      .lpszTitle = Title & Chr$(0)
   Else
      .lpszTitle = "选择目录" & Chr$(0)
    End If
End With
pidl = SHBrowseForFolder(bi)
If SHGetPathFromIDlist(ByVal pidl, ByVal folder) Then
    GetFolder = Left(folder, InStr(folder, Chr$(0)) - 1)
Else
    GetFolder = ""
End If
End Function

Public Function GetFileName(FilePathFileName As String) As String   '获取文件名  aaa.txt
    On Error Resume Next
    Dim i As Integer, J As Integer
    i = Len(FilePathFileName)
    J = InStrRev(FilePathFileName, "\")
    GetFileName = Mid(FilePathFileName, J + 1, i)
End Function

Public Function GetFilePath(sFile As String) As String
    '传入带全路径的字符串，返回文件路径
    On Error Resume Next
    Dim i As Long
    i = InStrRev(sFile, "\")
    GetFilePath = Left(sFile, i - 1)
End Function

