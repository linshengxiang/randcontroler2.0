VERSION 5.00
Begin VB.Form reg 
   Caption         =   "Form3"
   ClientHeight    =   4260
   ClientLeft      =   13710
   ClientTop       =   10785
   ClientWidth     =   4155
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   4155
   Begin VB.CommandButton Command3 
      Caption         =   "exit"
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   2055
      Left            =   720
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "reg.frx":0000
      Top             =   1560
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取消镜像劫持随机抽号"
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
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "镜像劫持随机抽号"
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
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "reg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo error1
Dim A
Set A = CreateObject("wscript.shell")
A.regwrite "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\随机抽号.exe\Debugger", "随机抽号2.0.exe", "REG_SZ"
Exit Sub

error1:
    MsgBox "error: 权限不足"
End Sub

Private Sub Command2_Click()
On Error GoTo error1
Dim A
Set A = CreateObject("wscript.shell")
A.regwrite "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\随机抽号.exe\Debugger", "", "REG_SZ"
Exit Sub

error1:
    MsgBox ("error: 权限不足")
End Sub

Private Sub Command3_Click()
Unload reg
End Sub
