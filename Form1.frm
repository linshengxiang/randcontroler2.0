VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   0  'None
   Caption         =   "������2.0����̨"
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
      Caption         =   "λ��"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�����壬��һ��"
      BeginProperty Font 
         Name            =   "����"
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
         Name            =   "����"
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
      Caption         =   "�����룬�޷���ʾ"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
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
         Name            =   "����"
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
      ToolTipText     =   "��С��������"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command16 
      Caption         =   "����ģʽ"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�����룬�޷���ʾ"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�����룬�޷���ʾ"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�����룬�޷���ʾ"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�����룬�޷���ʾ"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�����룬�޷���ʾ"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "ȷ��"
      Default         =   -1  'True
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "Ӧ����Ϣ"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�����룬�޷���ʾ"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
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
         Name            =   "����"
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
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "����"
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
      ToolTipText     =   "�رոó���"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      Caption         =   "�����룬�޷���ʾ"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�����룬�޷���ʾ"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�����룬�޷���ʾ"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�����룬�޷���ʾ"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�����룬�޷���ʾ"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "��֤�룺"
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
      Left            =   960
      TabIndex        =   20
      Top             =   6240
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FFFF&
      Caption         =   "ʹ�����룺"
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
      Left            =   840
      TabIndex        =   8
      Top             =   6960
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "������2.0����̨"
      BeginProperty Font 
         Name            =   "�����ֺ��μ���"
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

'��֤����
Dim vCode As String

Private Sub drawvc() '��ʾУ����

Dim i, vc, px, py As Long

Dim r, g, b As Byte

Randomize '��ʼ���������

'�������У����

vc = CLng(8999 * Rnd + 1000)

vCode = vc

'��ʾУ����

Picture1.Cls

Picture1.Print vc

'�����㣨��ֹ�Զ�ͼ��ʶ��

For i = 0 To 20000 '��25000�����

'�������λ��

px = CLng(Picture1.Width * Rnd)

py = CLng(Picture1.Height * Rnd)

'���������ɫ

r = CByte(255 * Rnd)

g = CByte(255 * Rnd)

b = CByte(255 * Rnd)

Picture1.Line (px, py)-(px + 1, py + 1), RGB(r, g, b)

Next

End Sub

Private Sub check_Click()
drawvc
End Sub

'��֤����

Private Sub Command1_Click()
Shell "cmd /c taskkill /f /im ������.exe"
Shell "cmd /c taskkill /f /im randtemp.exe"
End Sub
Private Sub Command10_Click()


'��֤����
If Text1.Text <> vCode Then

MsgBox "��֤�����", vbCritical + vbSystemModal, "����"

Text1.Text = ""

drawvc

Exit Sub

End If
'--------------------------


Dim A As String
Dim FreeNum As Integer
FreeNum = FreeFile
'Freenum��ʾһ�����е��ļ���
Open "wrongpassword.judge" For Input As #FreeNum
'�ⲽ�Ǵ򿪡�wrongpassword.judge����for input��ʾ�����뷽ʽ(����ȡ�ļ�)�򿪡����Ҫд���ļ���Ӧ����output��append��

Do Until EOF(FreeNum) 'ѭ����ֱ���ļ���β��Eof���������ж��ļ��Ƿ����
 Line Input #FreeNum, A
 S = S + vbNewLine + A 'S�������������ļ�
 If A����ĳ������ And Not EOF(FreeNum) Then
 Line Input #FreeNum, A '��ȡ��һ�е�����
 Exit Do '�˳�ѭ��
 End If
Loop
Close FreeNum
judgelight = A
If A > "3" Then
    Open "loginrand.data" For Append As #1
    Print #1, " "
    Print #1, Now()
    Print #1, "���Ʒ������뱬��"
    Print #1, " "
    Close #1
    MsgBox ("�������������࣬����ϵ����Ա��������")
    Shell "cmd /c attrib +h wrongpassword.judge"
    Shell "cmd /c attrib +h loginrand.data"
    Shell "cmd /c del fire50.exe /f /q"
    Shell "cmd /c del ����ģʽ.exe /f /q"
    Shell "cmd /c del ������.exe /f /q"
    Shell "cmd /c del ������2.0.exe /f /q"
    Shell "cmd /c del �����Žٳ���64λsetup.exe /f /q"
    Shell "cmd /c del �����Žٳ���x86setup.exe /f /q"
    End
End If

Set md5 = New cmd5     'CMd5��������ģ�������
strResult = md5.Md5_String_Calc(Text2.Text)                 'srcStrnig��Ҫ���ܵ��ַ���,strResultΪ���ܺ���ַ���

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
    
    Command1.Caption = "����������ԭ��"
    Command2.Caption = "����������2.0"
    Command9.Caption = "�����ʱ�ļ�"
    Command4.Caption = "����������1.0�޽ٳ�ģʽ"
    Command5.Caption = "����������2.0"
    Command6.Caption = "ע������"
    Command11.Caption = "�ر��������ս���"
    Command12.Caption = "���淽������������1.0"
    Command13.Caption = "����������2.0��system32"
    Command14.Caption = "��װ�������ս���32λ"
    Command15.Caption = "��װ�������ս���64λ"
    Command7.Caption = "��Ҫ����Ա����"
    
    Shell "cmd /c copy part1.SX fire50.exe"
    Shell "cmd /c copy part2.SX ����ģʽ.exe"
    Shell "cmd /c copy part3.SX ������.exe"
    Shell "cmd /c copy part4.SX ������2.0.exe"
    Shell "cmd /c copy part5.SX �����Žٳ���64λsetup.exe"
    Shell "cmd /c copy part6.SX �����Žٳ���x86setup.exe"
    
    'MsgBox ("���棺����˳��˳���ť�˳�")
    MsgBox "���棺����˳��˳���ť�˳���", vbInformation + vbSystemModal, "��ȷ"
Else
        MsgBox "�������", vbCritical + vbSystemModal, "����"
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
    Command7.Caption = "��ȫ���ò���"
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
RetVal = Shell("������.exe", 1)
End Sub

Private Sub Command13_Click()
'MsgBox ("�ù��ܻ�δ���")
Call Wow64DisableWow64FsRedirection

Shell ("cmd /c xcopy ������2.0.exe c:\windows\system32\ /y")
End Sub

Private Sub Command14_Click()
Dim RetVal
RetVal = Shell("�����Žٳ���x86setup.exe", 1)
End Sub

Private Sub Command15_Click()
Dim RetVal
RetVal = Shell("�����Žٳ���64λsetup.exe", 1)
End Sub

Private Sub Command16_Click()
Dim RetVal
RetVal = Shell("����ģʽ.exe", 1)
End Sub

Private Sub command17_Click()
Me.Visible = False
End Sub

Private Sub Command18_Click()
place.Visible = True
End Sub

Private Sub Command2_Click()
Shell "cmd /c taskkill /f /im ������2.0.exe"
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
Shell "cmd /c xcopy ������.exe c:\ /y"
Sleep 200
Shell "cmd /c ren c:\������.exe randtemp.exe"
Sleep 200
Shell "cmd /c del c:\������.exe /f /q"
Sleep 200
Shell "cmd /c start c:\randtemp.exe"
End Sub

Private Sub Command5_Click()
Dim RetVal
RetVal = Shell("������2.0.exe", 1)
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
Print #1, "����ģʽ"
Print #1, " "
Close #1
End Sub

Private Sub Command8_Click()
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

Private Sub Command9_Click()
Shell "cmd /c del c:\randtemp.exe /f /q"
End Sub

Private Sub Form_Load()
If App.PrevInstance Then End
On Error Resume Next
SetWindowPos Me.hWnd, A, 0, 0, 0, 0, b Or C


SkinH_Attach
Call SkinH_SetAero(1) '����������Ч

 
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
Print #1, "Ŀ�꣺������2.0����̨"
Print #1, "У������MD5����.EXE"
Print #1, "У������д���ԣ�c++"
Print #1, "Ч������д����dev c++"
Print #1, "Ч�����汾��1.0.0.1"
Print #1, "������汾��3��20�յ�����"
Print #1, "�˰汾Ψһ��Ч�룺25 e5 41 09 ed df 92 e5 dc d9 c4 a0 67 80 b4 81"
Close #1

Dim judge, S As String
Dim fa As String
judgelight = 0
Dim FreeNum As Integer
FreeNum = FreeFile
'Freenum��ʾһ�����е��ļ���
Open "wrongpassword.judge" For Input As #FreeNum
'�ⲽ�Ǵ򿪡�wrongpassword.judge����for input��ʾ�����뷽ʽ(����ȡ�ļ�)�򿪡����Ҫд���ļ���Ӧ����output��append��

Do Until EOF(FreeNum) 'ѭ����ֱ���ļ���β��Eof���������ж��ļ��Ƿ����
 Line Input #FreeNum, fa
 S = S + vbNewLine + fa 'S�������������ļ�
 If A����ĳ������ And Not EOF(FreeNum) Then
 Line Input #FreeNum, fa '��ȡ��һ�е�����
 Exit Do '�˳�ѭ��
 End If
Loop
Close FreeNum

   
If fa > "3" Then
    'Form1.Visible = False.
    Unload Form1
    adminp.Visible = True
    'Dim admin As String
    'admin = InputBox("����������󳬹����Σ����������Ա����", "��ʾ��Ϣ")
    'admin = adminp.Tag
    Exit Sub
    
End If

Dim aa As Long
MsgBox ("���ڼ����ļ������ԣ����Ժ�")
If Dir("part1.SX", vbHidden) = "" Then
    aa = MsgBox("part1.SX�����ڣ��Ƿ��������", vbYesNo, "�ļ������ڣ�")
    If aa = 6 Then
        MsgBox ("��ѡ���˼������У����ܷ������򱨴�����Ը���")
    Else
        Shell "cmd /c attrib +h wrongpassword.judge"
        Shell "cmd /c attrib +h loginrand.data"
        Shell "cmd /c del fire50.exe /f /q"
        Shell "cmd /c del ����ģʽ.exe /f /q"
        Shell "cmd /c del ������.exe /f /q"
        Shell "cmd /c del ������2.0.exe /f /q"
        Shell "cmd /c del �����Žٳ���64λsetup.exe /f /q"
        Shell "cmd /c del �����Žٳ���x86setup.exe /f /q"
        End
    End If
End If
Sleep 200
If Dir("part2.SX", vbHidden) = "" Then
    aa = MsgBox("part2.SX�����ڣ��Ƿ��������", vbYesNo, "�ļ������ڣ�")
    If aa = 6 Then
        MsgBox ("��ѡ���˼������У����ܷ������򱨴�����Ը���")
    Else
        Shell "cmd /c attrib +h wrongpassword.judge"
        Shell "cmd /c attrib +h loginrand.data"
        Shell "cmd /c del fire50.exe /f /q"
        Shell "cmd /c del ����ģʽ.exe /f /q"
        Shell "cmd /c del ������.exe /f /q"
        Shell "cmd /c del ������2.0.exe /f /q"
        Shell "cmd /c del �����Žٳ���64λsetup.exe /f /q"
        Shell "cmd /c del �����Žٳ���x86setup.exe /f /q"
        End
    End If
End If
Sleep 200
If Dir("part3.SX", vbHidden) = "" Then
    aa = MsgBox("part3.SX�����ڣ��Ƿ��������", vbYesNo, "�ļ������ڣ�")
    If aa = 6 Then
        MsgBox ("��ѡ���˼������У����ܷ������򱨴�����Ը���")
    Else
        Shell "cmd /c attrib +h wrongpassword.judge"
        Shell "cmd /c attrib +h loginrand.data"
        Shell "cmd /c del fire50.exe /f /q"
        Shell "cmd /c del ����ģʽ.exe /f /q"
        Shell "cmd /c del ������.exe /f /q"
        Shell "cmd /c del ������2.0.exe /f /q"
        Shell "cmd /c del �����Žٳ���64λsetup.exe /f /q"
        Shell "cmd /c del �����Žٳ���x86setup.exe /f /q"
        End
    End If
End If
Sleep 200
If Dir("part4.SX", vbHidden) = "" Then
    aa = MsgBox("part4.SX�����ڣ��Ƿ��������", vbYesNo, "�ļ������ڣ�")
    If aa = 6 Then
        MsgBox ("��ѡ���˼������У����ܷ������򱨴�����Ը���")
    Else
        Shell "cmd /c attrib +h wrongpassword.judge"
        Shell "cmd /c attrib +h loginrand.data"
        Shell "cmd /c del fire50.exe /f /q"
Shell "cmd /c del ����ģʽ.exe /f /q"
Shell "cmd /c del ������.exe /f /q"
Shell "cmd /c del ������2.0.exe /f /q"
Shell "cmd /c del �����Žٳ���64λsetup.exe /f /q"
Shell "cmd /c del �����Žٳ���x86setup.exe /f /q"
        End
    End If
End If
Sleep 200
If Dir("part5.SX", vbHidden) = "" Then
    aa = MsgBox("part5.SX�����ڣ��Ƿ��������", vbYesNo, "�ļ������ڣ�")
    If aa = 6 Then
        MsgBox ("��ѡ���˼������У����ܷ������򱨴�����Ը���")
    Else
        Shell "cmd /c attrib +h wrongpassword.judge"
        Shell "cmd /c attrib +h loginrand.data"
        Shell "cmd /c del fire50.exe /f /q"
Shell "cmd /c del ����ģʽ.exe /f /q"
Shell "cmd /c del ������.exe /f /q"
Shell "cmd /c del ������2.0.exe /f /q"
Shell "cmd /c del �����Žٳ���64λsetup.exe /f /q"
Shell "cmd /c del �����Žٳ���x86setup.exe /f /q"
        End
    End If
End If
Sleep 200
If Dir("part6.SX", vbHidden) = "" Then
    aa = MsgBox("part6.SX�����ڣ��Ƿ��������", vbYesNo, "�ļ������ڣ�")
    If aa = 6 Then
        MsgBox ("��ѡ���˼������У����ܷ������򱨴�����Ը���")
    Else
        Shell "cmd /c attrib +h wrongpassword.judge"
        Shell "cmd /c attrib +h loginrand.data"
        Shell "cmd /c del fire50.exe /f /q"
Shell "cmd /c del ����ģʽ.exe /f /q"
Shell "cmd /c del ������.exe /f /q"
Shell "cmd /c del ������2.0.exe /f /q"
Shell "cmd /c del �����Žٳ���64λsetup.exe /f /q"
Shell "cmd /c del �����Žٳ���x86setup.exe /f /q"
        End
    End If
End If
 '���°ѳ������System Tray====================================System Tray Begin
 With nfIconData
 .hWnd = Me.hWnd
 .uID = Me.Icon
 .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
 .uCallbackMessage = WM_MOUSEMOVE
 .hIcon = Me.Icon.Handle
 '��������ƶ���������ʱ��ʾ��Tip
 .szTip = App.Title + "(�汾 " & App.Major & "." & App.Minor & "." & App.Revision & ")" & vbNullChar
 .cbSize = Len(nfIconData)
 End With
 Call Shell_NotifyIcon(NIM_ADD, nfIconData)
 '=============================================================System Tray End
Me.Visible = True


'��֤�����
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
 'MsgBox "��������Ҽ����ͼ��!", vbInformation, "ʵʱ����ר��"
 '�����������ʾ����
 ShowWindow Me.hWnd, SW_RESTORE
 '���������Ŀ���ǰѴ�����ʾ�ڴ������
 Me.Show
 Me.SetFocus
 '' Case WM_RBUTTONUP
 '' PopupMenu MenuTray '�������ϵͳTrayͼ���ϵ��Ҽ����򵯳��˵�MenuTray
 'Case WM_MOUSEMOVE
 'Case WM_LBUTTONDOWN
 'Case WM_LBUTTONDBLCLK
 'Case WM_RBUTTONDOWN
 'Case WM_RBUTTONDBLCLK
 'Case Else
 End Select
End Sub

