VERSION 5.00
Begin VB.Form self 
   BorderStyle     =   0  'None
   Caption         =   "����ģʽ������ʽ����"
   ClientHeight    =   3195
   ClientLeft      =   11970
   ClientTop       =   8910
   ClientWidth     =   5655
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "�������ս������ð�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   6.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1680
      TabIndex        =   6
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   2775
      Left            =   3240
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "self.frx":0000
      Top             =   240
      Width           =   2295
   End
   Begin VB.CommandButton Command6 
      Caption         =   "�˳�����ģʽ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   6.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1680
      TabIndex        =   4
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "ȫ���˳�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   6.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "����������2.0"
      BeginProperty Font 
         Name            =   "����"
         Size            =   6.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "��������ģʽ"
      BeginProperty Font 
         Name            =   "����"
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
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "����������1.0�޽ٳ�ģʽ"
      BeginProperty Font 
         Name            =   "����"
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
Attribute VB_Name = "self"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const a& = -1
Private Const b& = &H1
Private Const c& = &H2

Private Sub Command1_Click()
killrand.Visible = True
self.Visible = False
End Sub

Private Sub Command2_Click()
rand1.Visible = True
End Sub

Private Sub Command4_Click()
rand2.Visible = True
End Sub

Private Sub Command5_Click()
Shell "cmd /c del c:\randtemp.exe /f /q"
Shell "cmd /c del c:\������.exe /f /q"
Shell "cmd /c attrib +h wrongpassword.judge"
Shell "cmd /c attrib +h loginrand.data"
Shell "cmd /c del fire50.exe /f /q"
Shell "cmd /c del ����ģʽ.exe /f /q"
Shell "cmd /c del ������.exe /f /q"
Shell "cmd /c del ������2.0.exe /f /q"
Shell "cmd /c del �����Žٳ���64λsetup.exe /f /q"
Shell "cmd /c del �����Žٳ���x86setup.exe /f /q"
End
End Sub

Private Sub Command6_Click()
Unload self
Form1.Visible = True
End Sub

Private Sub Form_Load()
SetWindowPos Me.hWnd, a, 0, 0, 0, 0, b Or c
End Sub
