VERSION 5.00
Begin VB.Form adminp 
   Caption         =   "密码错误次数过多"
   ClientHeight    =   3030
   ClientLeft      =   12090
   ClientTop       =   5925
   ClientWidth     =   5670
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   5670
   Begin VB.CommandButton Command2 
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
      Height          =   615
      Left            =   2880
      TabIndex        =   7
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确认"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   6
      Top             =   2280
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   4200
      ScaleHeight     =   795
      ScaleWidth      =   915
      TabIndex        =   4
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1440
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
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
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "由于密码错误次数过多，请输入管理员密码解锁"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Top             =   240
      Width           =   4575
   End
   Begin VB.Label Label2 
      Caption         =   "验证码："
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "管理员密码："
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   2280
      Width           =   1575
   End
End
Attribute VB_Name = "adminp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

For i = 0 To 10000 '画10000个噪点

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



Private Sub Command1_Click()

If Text2.Text <> vCode Then

MsgBox "验证码错误。", vbCritical + vbSystemModal, "错误"

Text2.Text = ""

drawvc

End If

admin = Text1.Text

Set md5 = New cmd5     'CMd5是新增类模块的名称
    strResult = md5.Md5_String_Calc(admin)
    If strResult <> "89FA34279C9E4F5B82E76A1EB0FFF075" Then
        MsgBox ("Error password!!")
        Open "loginrand.data" For Append As #1
        Print #1, " "
        Print #1, Now()
        Print #1, "发现爆破管理员密码"
        Print #1, " "
        Close #1
        Shell "cmd /c attrib +h wrongpassword.judge"
        Shell "cmd /c attrib +h loginrand.data"
        Shell "cmd /c del fire50.exe /f /q"
        Shell "cmd /c del 命令模式.exe /f /q"
        Shell "cmd /c del 随机抽号.exe /f /q"
        Shell "cmd /c del 随机抽号2.0.exe /f /q"
        Shell "cmd /c del 随机抽号劫持器64位setup.exe /f /q"
        Shell "cmd /c del 随机抽号劫持器x86setup.exe /f /q"
        End
    Else
        Open "wrongpassword.judge" For Output As #1
        Print #1, judgelight
        Close #1
        Form1.Visible = True
        Unload adminp
    End If
End Sub

Private Sub Command2_Click()
drawvc
End Sub

Private Sub Form_Load()

Picture1.FontSize = 12

Picture1.FontBold = True

Picture1.AutoRedraw = True

drawvc

End Sub


