VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   0  'None
   Caption         =   "随机抽号2.0控制台"
   ClientHeight    =   7545
   ClientLeft      =   5280
   ClientTop       =   5265
   ClientWidth     =   6960
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command18 
      Caption         =   "位置"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   6.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   24
      Top             =   480
      Width           =   735
   End
   Begin VB.CommandButton check 
      Caption         =   "看不清，换一张"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   23
      Top             =   6120
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   735
      Left            =   5040
      ScaleHeight     =   675
      ScaleWidth      =   1395
      TabIndex        =   22
      Top             =   6000
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   21
      Top             =   6120
      Width           =   1575
   End
   Begin VB.CommandButton Command7 
      Caption         =   "无密码，无法显示"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4560
      TabIndex        =   19
      Top             =   4920
      Width           =   2055
   End
   Begin VB.CommandButton command17 
      Caption         =   "--"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   5.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      TabIndex        =   18
      ToolTipText     =   "最小化到托盘"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command16 
      Caption         =   "命令模式"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   5.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      MaskColor       =   &H00808080&
      TabIndex        =   17
      Top             =   6360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command15 
      Caption         =   "无密码，无法显示"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4560
      TabIndex        =   16
      Top             =   3600
      Width           =   2055
   End
   Begin VB.CommandButton Command14 
      Caption         =   "无密码，无法显示"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4560
      TabIndex        =   15
      Top             =   2280
      Width           =   2055
   End
   Begin VB.CommandButton Command13 
      Caption         =   "无密码，无法显示"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4560
      TabIndex        =   14
      Top             =   960
      Width           =   2055
   End
   Begin VB.CommandButton Command12 
      Caption         =   "无密码，无法显示"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2040
      TabIndex        =   13
      Top             =   4920
      Width           =   2415
   End
   Begin VB.CommandButton Command11 
      Caption         =   "无密码，无法显示"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   12
      Top             =   4920
      Width           =   1815
   End
   Begin VB.CommandButton Command10 
      Caption         =   "确认"
      Default         =   -1  'True
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   11
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "应用信息"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   6.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   10
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton Command9 
      Caption         =   "无密码，无法显示"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   9
      Top             =   3600
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   6840
      Width           =   3015
   End
   Begin VB.CommandButton Command8 
      Caption         =   "×"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   6.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6000
      TabIndex        =   6
      ToolTipText     =   "关闭该程序"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      Caption         =   "无密码，无法显示"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2040
      TabIndex        =   5
      Top             =   3600
      Width           =   2415
   End
   Begin VB.CommandButton Command5 
      Caption         =   "无密码，无法显示"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2040
      TabIndex        =   4
      Top             =   2280
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      Caption         =   "无密码，无法显示"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2040
      TabIndex        =   3
      Top             =   960
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "无密码，无法显示"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   6.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "无密码，无法显示"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FFFF&
      Caption         =   "验证码："
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
      Left            =   960
      TabIndex        =   20
      Top             =   6240
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FFFF&
      Caption         =   "使用密码："
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
      Left            =   840
      TabIndex        =   8
      Top             =   6960
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "随机抽号2.0控制台"
      BeginProperty Font 
         Name            =   "方正粗黑宋简体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const A& = -1
Private Const b& = &H1
Private Const C& = &H2

Private Declare Sub Sleep Lib "kernel32.DLL" (ByVal dwMilliseconds As Long)

Private Declare Function Wow64RevertWow64FsRedirection Lib "kernel32.DLL" ()
Private Declare Function Wow64DisableWow64FsRedirection Lib "kernel32.DLL" ()

Dim judgelight As String

'验证码检查
Dim vCode As String

Private Sub drawvc() '显示校验码

Dim i, vc, px, py As Long

Dim r, g, b As Byte

Randomize '初始化随机种子

'生成随机校验码

vc = CLng(8999 * Rnd + 1000)

vCode = vc

'显示校验码

Picture1.Cls

Picture1.Print vc

'添加噪点（防止自动图像识别）

For i = 0 To 20000 '画25000个噪点

'画点随机位置

px = CLng(Picture1.Width * Rnd)

py = CLng(Picture1.Height * Rnd)

'画点随机颜色

r = CByte(255 * Rnd)

g = CByte(255 * Rnd)

b = CByte(255 * Rnd)

Picture1.Line (px, py)-(px + 1, py + 1), RGB(r, g, b)

Next

End Sub

Private Sub check_Click()
drawvc
End Sub

'验证码检查

Private Sub Command1_Click()
Shell "cmd /c taskkill /f /im 随机抽号.exe"
Shell "cmd /c taskkill /f /im randtemp.exe"
End Sub
Private Sub Command10_Click()


'验证码检查
If Text1.Text <> vCode Then

MsgBox "验证码错误。", vbCritical + vbSystemModal, "错误"

Text1.Text = ""

drawvc

Exit Sub

End If
'--------------------------


Dim A As String
Dim FreeNum As Integer
FreeNum = FreeFile
'Freenum表示一个空闲的文件号
Open "wrongpassword.judge" For Input As #FreeNum
'这步是打开“wrongpassword.judge”，for input表示以输入方式(即读取文件)打开。如果要写入文件则应该用output或append。

Do Until EOF(FreeNum) '循环，直到文件结尾。Eof函数用来判断文件是否读完
 Line Input #FreeNum, A
 S = S + vbNewLine + A 'S用来保存整个文件
 If A满足某个条件 And Not EOF(FreeNum) Then
 Line Input #FreeNum, A '读取下一行的内容
 Exit Do '退出循环
 End If
Loop
Close FreeNum
judgelight = A
If A > "3" Then
    Open "loginrand.data" For Append As #1
    Print #1, " "
    Print #1, Now()
    Print #1, "疑似发生密码爆破"
    Print #1, " "
    Close #1
    MsgBox ("密码错误次数过多，请联系管理员解锁！！")
    Shell "cmd /c attrib +h wrongpassword.judge"
    Shell "cmd /c attrib +h loginrand.data"
    Shell "cmd /c del fire50.exe /f /q"
    Shell "cmd /c del 命令模式.exe /f /q"
    Shell "cmd /c del 随机抽号.exe /f /q"
    Shell "cmd /c del 随机抽号2.0.exe /f /q"
    Shell "cmd /c del 随机抽号劫持器64位setup.exe /f /q"
    Shell "cmd /c del 随机抽号劫持器x86setup.exe /f /q"
    End
End If

Set md5 = New cmd5     'CMd5是新增类模块的名称
strResult = md5.Md5_String_Calc(Text2.Text)                 'srcStrnig是要加密的字符串,strResult为加密后的字符串

If strResult = "8C58FD74056C6AEF4A0AFF52468262BA" Or strResult = "89FA34279C9E4F5B82E76A1EB0FFF075" Then
    Open "wrongpassword.judge" For Output As #1
    Print #1, "0"
    Close #1
    
    Open "loginrand.data" For Append As #1
    Print #1, Now()
    Print #1, "passwordright"
    Print #1, " "
    Close #1
    
    Command1.Enabled = True
    Command2.Enabled = True
    Command9.Enabled = True
    Command4.Enabled = True
    Command5.Enabled = True
    Command6.Enabled = True
    Command11.Enabled = True
    Command12.Enabled = True
    Command13.Enabled = True
    Command14.Enabled = True
    Command15.Enabled = True
    'Command7.Enabled = True
    
    Command1.Caption = "结束随机抽号原版"
    Command2.Caption = "结束随机抽号2.0"
    Command9.Caption = "清除临时文件"
    Command4.Caption = "启动随机抽号1.0无劫持模式"
    Command5.Caption = "启动随机抽号2.0"
    Command6.Caption = "注册表操作"
    Command11.Caption = "关闭随机抽号终结者"
    Command12.Caption = "常规方法启动随机抽号1.0"
    Command13.Caption = "复制随机抽号2.0到system32"
    Command14.Caption = "安装随机抽号终结者32位"
    Command15.Caption = "安装随机抽号终结者64位"
    Command7.Caption = "需要管理员密码"
    
    Shell "cmd /c copy part1.SX fire50.exe"
    Shell "cmd /c copy part2.SX 命令模式.exe"
    Shell "cmd /c copy part3.SX 随机抽号.exe"
    Shell "cmd /c copy part4.SX 随机抽号2.0.exe"
    Shell "cmd /c copy part5.SX 随机抽号劫持器64位setup.exe"
    Shell "cmd /c copy part6.SX 随机抽号劫持器x86setup.exe"
    
    'MsgBox ("警告：请从退出此程序按钮退出")
    MsgBox "警告：请从退出此程序按钮退出。", vbInformation + vbSystemModal, "正确"
Else
        MsgBox "密码错误。", vbCritical + vbSystemModal, "错误"
        judgelight = judgelight + 1
        Open "wrongpassword.judge" For Output As #1
        Print #1, judgelight
        Close #1
        
        Open "loginrand.data" For Append As #1
        Print #1, " "
        Print #1, Now()
        Print #1, "wrongpassword"
        Close #1
        drawvc
End If

If strResult = "89FA34279C9E4F5B82E76A1EB0FFF075" Then
    Command16.Visible = True
    Command7.Caption = "完全内置操作"
    Command7.Enabled = True
End If
Text2.Text = ""
Text1.Text = ""
End Sub

Private Sub Command11_Click()
Shell "cmd /c taskkill /f /im cmd.exe"
End Sub

Private Sub Command12_Click()
Dim RetVal
RetVal = Shell("随机抽号.exe", 1)
End Sub

Private Sub Command13_Click()
'MsgBox ("该功能还未完成")
Call Wow64DisableWow64FsRedirection

Shell ("cmd /c xcopy 随机抽号2.0.exe c:\windows\system32\ /y")
End Sub

Private Sub Command14_Click()
Dim RetVal
RetVal = Shell("随机抽号劫持器x86setup.exe", 1)
End Sub

Private Sub Command15_Click()
Dim RetVal
RetVal = Shell("随机抽号劫持器64位setup.exe", 1)
End Sub

Private Sub Command16_Click()
Dim RetVal
RetVal = Shell("命令模式.exe", 1)
End Sub

Private Sub command17_Click()
Me.Visible = False
End Sub

Private Sub Command18_Click()
place.Visible = True
End Sub

Private Sub Command2_Click()
Shell "cmd /c taskkill /f /im 随机抽号2.0.exe"
End Sub

Private Sub Command3_Click()
Shell "MD5check.exe"
Sleep 300
Form2.Visible = True
Form1.Visible = False
Command10.Enabled = True
'Text2.Locked = False
End Sub

Private Sub Command4_Click()
Shell "cmd /c xcopy 随机抽号.exe c:\ /y"
Sleep 200
Shell "cmd /c ren c:\随机抽号.exe randtemp.exe"
Sleep 200
Shell "cmd /c del c:\随机抽号.exe /f /q"
Sleep 200
Shell "cmd /c start c:\randtemp.exe"
End Sub

Private Sub Command5_Click()
Dim RetVal
RetVal = Shell("随机抽号2.0.exe", 1)
End Sub

Private Sub Command6_Click()
reg.Visible = True
End Sub

Private Sub Command7_Click()
self.Visible = True
Form1.Visible = False
Open "loginrand.data" For Append As #1
Print #1, " "
Print #1, Now()
Print #1, "内置模式"
Print #1, " "
Close #1
End Sub

Private Sub Command8_Click()
Shell "cmd /c del c:\randtemp.exe /f /q"
Shell "cmd /c del c:\随机抽号.exe /f /q"
Shell "cmd /c attrib +h wrongpassword.judge"
Shell "cmd /c attrib +h loginrand.data"
Shell "cmd /c del fire50.exe /f /q"
Shell "cmd /c del 命令模式.exe /f /q"
Shell "cmd /c del 随机抽号.exe /f /q"
Shell "cmd /c del 随机抽号2.0.exe /f /q"
Shell "cmd /c del 随机抽号劫持器64位setup.exe /f /q"
Shell "cmd /c del 随机抽号劫持器x86setup.exe /f /q"
End
End Sub

Private Sub Command9_Click()
Shell "cmd /c del c:\randtemp.exe /f /q"
End Sub

Private Sub Form_Load()
If App.PrevInstance Then End
On Error Resume Next
SetWindowPos Me.hWnd, A, 0, 0, 0, 0, b Or C


SkinH_Attach
Call SkinH_SetAero(1) '开启窗体特效

 
If Dir("wrongpassword.judge", vbHidden) = "" Then
    Open "wrongpassword.judge" For Output As #1
    Print #1, "0"
    Close #1
End If
If Dir("loginrand.data", vbHidden) = "" Then
    Open "loginrand.data" For Output As #1
    Print #1, ""
    Close #1
End If

Shell "cmd /c attrib -h wrongpassword.judge"
Shell "cmd /c attrib -h loginrand.data"

Open "loginrand.data" For Append As #1
Print #1, Now()
Print #1, "open"
Print #1, " "
Close #1

Open "md5fc.jud" For Output As #1
Print #1, "目标：随机抽号2.0控制台"
Print #1, "校验器：MD5加密.EXE"
Print #1, "校验器编写语言：c++"
Print #1, "效验器编写程序：dev c++"
Print #1, "效验器版本：1.0.0.1"
Print #1, "本程序版本：3月20日第三代"
Print #1, "此版本唯一有效码：25 e5 41 09 ed df 92 e5 dc d9 c4 a0 67 80 b4 81"
Close #1

Dim judge, S As String
Dim fa As String
judgelight = 0
Dim FreeNum As Integer
FreeNum = FreeFile
'Freenum表示一个空闲的文件号
Open "wrongpassword.judge" For Input As #FreeNum
'这步是打开“wrongpassword.judge”，for input表示以输入方式(即读取文件)打开。如果要写入文件则应该用output或append。

Do Until EOF(FreeNum) '循环，直到文件结尾。Eof函数用来判断文件是否读完
 Line Input #FreeNum, fa
 S = S + vbNewLine + fa 'S用来保存整个文件
 If A满足某个条件 And Not EOF(FreeNum) Then
 Line Input #FreeNum, fa '读取下一行的内容
 Exit Do '退出循环
 End If
Loop
Close FreeNum

   
If fa > "3" Then
    'Form1.Visible = False.
    Unload Form1
    adminp.Visible = True
    'Dim admin As String
    'admin = InputBox("由于密码错误超过三次，请输入管理员密码", "提示信息")
    'admin = adminp.Tag
    Exit Sub
    
End If

Dim aa As Long
MsgBox ("正在检验文件完整性，请稍后。")
If Dir("part1.SX", vbHidden) = "" Then
    aa = MsgBox("part1.SX不存在，是否继续运行", vbYesNo, "文件不存在！")
    If aa = 6 Then
        MsgBox ("你选择了继续运行，可能发生程序报错，后果自负！")
    Else
        Shell "cmd /c attrib +h wrongpassword.judge"
        Shell "cmd /c attrib +h loginrand.data"
        Shell "cmd /c del fire50.exe /f /q"
        Shell "cmd /c del 命令模式.exe /f /q"
        Shell "cmd /c del 随机抽号.exe /f /q"
        Shell "cmd /c del 随机抽号2.0.exe /f /q"
        Shell "cmd /c del 随机抽号劫持器64位setup.exe /f /q"
        Shell "cmd /c del 随机抽号劫持器x86setup.exe /f /q"
        End
    End If
End If
Sleep 200
If Dir("part2.SX", vbHidden) = "" Then
    aa = MsgBox("part2.SX不存在，是否继续运行", vbYesNo, "文件不存在！")
    If aa = 6 Then
        MsgBox ("你选择了继续运行，可能发生程序报错，后果自负！")
    Else
        Shell "cmd /c attrib +h wrongpassword.judge"
        Shell "cmd /c attrib +h loginrand.data"
        Shell "cmd /c del fire50.exe /f /q"
        Shell "cmd /c del 命令模式.exe /f /q"
        Shell "cmd /c del 随机抽号.exe /f /q"
        Shell "cmd /c del 随机抽号2.0.exe /f /q"
        Shell "cmd /c del 随机抽号劫持器64位setup.exe /f /q"
        Shell "cmd /c del 随机抽号劫持器x86setup.exe /f /q"
        End
    End If
End If
Sleep 200
If Dir("part3.SX", vbHidden) = "" Then
    aa = MsgBox("part3.SX不存在，是否继续运行", vbYesNo, "文件不存在！")
    If aa = 6 Then
        MsgBox ("你选择了继续运行，可能发生程序报错，后果自负！")
    Else
        Shell "cmd /c attrib +h wrongpassword.judge"
        Shell "cmd /c attrib +h loginrand.data"
        Shell "cmd /c del fire50.exe /f /q"
        Shell "cmd /c del 命令模式.exe /f /q"
        Shell "cmd /c del 随机抽号.exe /f /q"
        Shell "cmd /c del 随机抽号2.0.exe /f /q"
        Shell "cmd /c del 随机抽号劫持器64位setup.exe /f /q"
        Shell "cmd /c del 随机抽号劫持器x86setup.exe /f /q"
        End
    End If
End If
Sleep 200
If Dir("part4.SX", vbHidden) = "" Then
    aa = MsgBox("part4.SX不存在，是否继续运行", vbYesNo, "文件不存在！")
    If aa = 6 Then
        MsgBox ("你选择了继续运行，可能发生程序报错，后果自负！")
    Else
        Shell "cmd /c attrib +h wrongpassword.judge"
        Shell "cmd /c attrib +h loginrand.data"
        Shell "cmd /c del fire50.exe /f /q"
Shell "cmd /c del 命令模式.exe /f /q"
Shell "cmd /c del 随机抽号.exe /f /q"
Shell "cmd /c del 随机抽号2.0.exe /f /q"
Shell "cmd /c del 随机抽号劫持器64位setup.exe /f /q"
Shell "cmd /c del 随机抽号劫持器x86setup.exe /f /q"
        End
    End If
End If
Sleep 200
If Dir("part5.SX", vbHidden) = "" Then
    aa = MsgBox("part5.SX不存在，是否继续运行", vbYesNo, "文件不存在！")
    If aa = 6 Then
        MsgBox ("你选择了继续运行，可能发生程序报错，后果自负！")
    Else
        Shell "cmd /c attrib +h wrongpassword.judge"
        Shell "cmd /c attrib +h loginrand.data"
        Shell "cmd /c del fire50.exe /f /q"
Shell "cmd /c del 命令模式.exe /f /q"
Shell "cmd /c del 随机抽号.exe /f /q"
Shell "cmd /c del 随机抽号2.0.exe /f /q"
Shell "cmd /c del 随机抽号劫持器64位setup.exe /f /q"
Shell "cmd /c del 随机抽号劫持器x86setup.exe /f /q"
        End
    End If
End If
Sleep 200
If Dir("part6.SX", vbHidden) = "" Then
    aa = MsgBox("part6.SX不存在，是否继续运行", vbYesNo, "文件不存在！")
    If aa = 6 Then
        MsgBox ("你选择了继续运行，可能发生程序报错，后果自负！")
    Else
        Shell "cmd /c attrib +h wrongpassword.judge"
        Shell "cmd /c attrib +h loginrand.data"
        Shell "cmd /c del fire50.exe /f /q"
Shell "cmd /c del 命令模式.exe /f /q"
Shell "cmd /c del 随机抽号.exe /f /q"
Shell "cmd /c del 随机抽号2.0.exe /f /q"
Shell "cmd /c del 随机抽号劫持器64位setup.exe /f /q"
Shell "cmd /c del 随机抽号劫持器x86setup.exe /f /q"
        End
    End If
End If
 '以下把程序放入System Tray====================================System Tray Begin
 With nfIconData
 .hWnd = Me.hWnd
 .uID = Me.Icon
 .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
 .uCallbackMessage = WM_MOUSEMOVE
 .hIcon = Me.Icon.Handle
 '定义鼠标移动到托盘上时显示的Tip
 .szTip = App.Title + "(版本 " & App.Major & "." & App.Minor & "." & App.Revision & ")" & vbNullChar
 .cbSize = Len(nfIconData)
 End With
 Call Shell_NotifyIcon(NIM_ADD, nfIconData)
 '=============================================================System Tray End
Me.Visible = True


'验证码操作
Picture1.FontSize = 12

Picture1.FontBold = True

Picture1.AutoRedraw = True

drawvc
'----------------------------------

End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Call Shell_NotifyIcon(NIM_DELETE, nfIconData)
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
 Dim lMsg As Single
 lMsg = X / Screen.TwipsPerPixelX
 Select Case lMsg
 Case WM_LBUTTONUP
 'MsgBox "请用鼠标右键点击图标!", vbInformation, "实时播音专家"
 '单击左键，显示窗体
 ShowWindow Me.hWnd, SW_RESTORE
 '下面两句的目的是把窗口显示在窗口最顶层
 Me.Show
 Me.SetFocus
 '' Case WM_RBUTTONUP
 '' PopupMenu MenuTray '如果是在系统Tray图标上点右键，则弹出菜单MenuTray
 'Case WM_MOUSEMOVE
 'Case WM_LBUTTONDOWN
 'Case WM_LBUTTONDBLCLK
 'Case WM_RBUTTONDOWN
 'Case WM_RBUTTONDBLCLK
 'Case Else
 End Select
End Sub

