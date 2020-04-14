VERSION 5.00
Begin VB.Form killrand 
   BorderStyle     =   0  'None
   Caption         =   "随机抽号终结者内置版控制台"
   ClientHeight    =   3015
   ClientLeft      =   12780
   ClientTop       =   10935
   ClientWidth     =   3540
   LinkTopic       =   "Form3"
   ScaleHeight     =   3015
   ScaleWidth      =   3540
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command6 
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   6.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   5
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "停止随机抽号终结者内置版"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   6.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1680
      TabIndex        =   4
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "隐藏"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   6.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "结束随机抽号1.0"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   6.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "启动随机抽号终结者内置版"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   6.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1680
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "扫描随机抽号1.0是否存在"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   6.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "killrand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const a& = -1
Private Const b& = &H1
Private Const c& = &H2

Private Declare Sub Sleep Lib "kernel32.DLL" (ByVal dwMilliseconds As Long)
Dim kk As Boolean
Private Sub Command1_Click()
If CheckExeIsRun("随机抽号.exe") Or rand1.Visible = True Then
    Dim rel
    rel = MsgBox("发现随机抽号.exe，是否强制结束", vbYesNo)
    If rel = 6 Then
        On Error Resume Next
        Dim s
        s = "随机抽号.exe"
        Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
        Set colProcessList = objWMIService.ExecQuery _
        ("Select * from Win32_Process Where Name='" & s & "'")
        For Each objProcess In colProcessList
        'MsgBox "已发现目标!"
        objProcess.Terminate '结束进程
        Next
        Set objProcess = Nothing
        Set colProcessList = Nothing
        Set objWMIService = Nothing
        rand1.Visible = False
    End If
Else
    Dim rev
    rev = MsgBox("不存在，是否打开", vbYesNo)
    If rev = 6 Then
        rand1.Visible = True
    End If
End If
End Sub
'检查进程是否运行，exeName 参数是要检查的进程 exe 名字，比如 VB6.EXE
Private Function CheckExeIsRun(exeName As String) As Boolean
    On Error GoTo Err
    Dim WMI
    Dim Obj
    Dim Objs
    CheckExeIsRun = False
    Set WMI = GetObject("WinMgmts:")
    Set Objs = WMI.InstancesOf("Win32_Process")
    For Each Obj In Objs
      If (InStr(UCase(exeName), UCase(Obj.Description)) <> 0) Then
            CheckExeIsRun = True
            If Not Objs Is Nothing Then Set Objs = Nothing
            If Not WMI Is Nothing Then Set WMI = Nothing
            Exit Function
      End If
    Next
    If Not Objs Is Nothing Then Set Objs = Nothing
    If Not WMI Is Nothing Then Set WMI = Nothing
    Exit Function
Err:
    If Not Objs Is Nothing Then Set Objs = Nothing
    If Not WMI Is Nothing Then Set WMI = Nothing
End Function

Private Sub Command2_Click()
kk = False
While kk = False
    If CheckExeIsRun("随机抽号.exe") Then
    On Error Resume Next
    Dim s
    s = "随机抽号.exe"
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set colProcessList = objWMIService.ExecQuery _
    ("Select * from Win32_Process Where Name='" & s & "'")
    For Each objProcess In colProcessList
    'MsgBox "已发现目标!"
    objProcess.Terminate '结束进程
    Next
    Set objProcess = Nothing
    Set colProcessList = Nothing
    Set objWMIService = Nothing
    rand2.Visible = True
    End If
    DoEvents
    'Sleep 3
Wend
End Sub

Private Sub Command3_Click()
On Error Resume Next
Dim s
s = "随机抽号.exe"
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set colProcessList = objWMIService.ExecQuery _
("Select * from Win32_Process Where Name='" & s & "'")
For Each objProcess In colProcessList
'MsgBox "已发现目标!"
objProcess.Terminate '结束进程
Next
Set objProcess = Nothing
Set colProcessList = Nothing
Set objWMIService = Nothing
End Sub

Private Sub Command4_Click()
killrand.Visible = False
End Sub

Private Sub Command5_Click()
kk = True
End Sub

Private Sub Command6_Click()
kk = True
Unload killrand
self.Visible = True
End Sub

Private Sub Form_Load()
SetWindowPos Me.hWnd, a, 0, 0, 0, 0, b Or c
End Sub
