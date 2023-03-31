VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmMain 
   Caption         =   "Flo2dAutomator"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4485
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4485
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   2
      Left            =   480
      Top             =   2640
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3840
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdFlo2d 
      Caption         =   "导入"
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton CmdMdl 
      Caption         =   "导入"
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox TxtMdl 
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Text            =   "请导入模型工程文件夹"
      Top             =   720
      Width           =   3015
   End
   Begin VB.TextBox TxtExe 
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Text            =   "请指定Flo2D驱动程序"
      Top             =   1440
      Width           =   3015
   End
   Begin VB.CommandButton ButtonExe 
      Caption         =   "运行模型"
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Flo2D可执行文件："
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "模型工程文件夹："
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   480
      Width           =   1575
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const NORMAL_PRIORITY_CLASS As Long = &H20&
Private Const STARTF_USESTDHANDLES As Long = &H100&
Private Const STARTF_USESHOWWINDOW As Long = &H1&
Private Const SW_HIDE As Long = 0&
Private Const SW_NORMAL As Long = 1
Private Const INFINITE As Long = &HFFFF&

Private Type PROCESS_INFORMATION
        hProcess As Long
        hThread As Long
        dwProcessId As Long
        dwThreadId As Long
End Type

Private Type STARTUPINFO
        cb As Long
        lpReserved As String
        lpDesktop As String
        lpTitle As String
        dwX As Long
        dwY As Long
        dwXSize As Long
        dwYSize As Long
        dwXCountChars As Long
        dwYCountChars As Long
        dwFillAttribute As Long
        dwFlags As Long
        wShowWindow As Integer
        cbReserved2 As Integer
        lpReserved2 As Long
        hStdInput As Long
        hStdOutput As Long
        hStdError As Long

End Type


Private Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long

End Type

Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long



Private Sub ButtonExe_Click()
    Dim si As STARTUPINFO
    Dim pi As PROCESS_INFORMATION
    Dim sa As SECURITY_ATTRIBUTES

    With si
    .cb = Len(si)
    .dwFlags = STARTF_USESTDHANDLES Or STARTF_USESHOWWINDOW
    .wShowWindow = SW_HIDE
    End With
    
    DoEvents
    CreateProcess TxtExe.Text, "", sa, sa, 0, 0, 0, TxtMdl.Text, si, pi
    FrmMain.Timer1.Enabled = True
End Sub

Private Sub CmdFlo2d_Click()
    Dim modelpath As String
    CommonDialog1.ShowOpen
    modelpath = CommonDialog1.FileName
    If modelpath <> "" Then
        TxtExe.Text = modelpath
    End If
End Sub

Private Sub CmdMdl_Click()
    Dim path As String
    path = GetFolder(Me.hwnd, "浏览文件夹")
    If path <> "" Then
        TxtMdl.Text = path
    End If
End Sub


Private Sub Form_Load()
    Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()
    Dim hWindSum As String
    Const WM_CLOSE = &H10
    hWindSum = FindWindow("#32770", "Simulation Summary")
    SendMessage hWindSum, WM_CLOSE, 0, 0
    'FrmMain.Timer1.Enabled = False

End Sub
