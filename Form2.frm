VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   4530
   ClientLeft      =   6285
   ClientTop       =   8715
   ClientWidth     =   4170
   LinkTopic       =   "Form2"
   ScaleHeight     =   4530
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Check1 
      Caption         =   "�ҽ��ܴ����"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   4080
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ��"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Top             =   3960
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   3615
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form2.frx":0000
      Top             =   240
      Width           =   3375
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const a& = -1
Private Const b& = &H1
Private Const c& = &H2
Private Declare Sub Sleep Lib "kernel32.DLL" (ByVal dwMilliseconds As Long)
Private Sub Check1_Click()
If Check1.Value = 1 Then
Check1.ForeColor = 0
Command1.Enabled = True
Else
Check1.ForeColor = &H80000011
Command1.Enabled = False
End If
End Sub

Private Sub Command1_Click()
Form1.Visible = True
Unload Form2
End Sub

Private Sub Form_Load()
SetWindowPos Me.hWnd, a, 0, 0, 0, 0, b Or c



Sleep 20

Dim judge, s As String
Dim fa As String
judgelight = 0
Dim FreeNum As Integer
FreeNum = FreeFile
'Freenum��ʾһ�����е��ļ���
Open "md5fc.log" For Input As #FreeNum
'�ⲽ�Ǵ򿪡�md5fc.log����for input��uʾ�����뷽ʽ(����ȡ�ļ�)�򿪡����Ҫд���ļ���Ӧ����output��append��

Do Until EOF(FreeNum) 'ѭ����ֱ���ļ���β��Eof���������ж��ļ��Ƿ����
 Line Input #FreeNum, fa
 s = s + vbNewLine + fa 'S�������������ļ�
 If A����ĳ������ And Not EOF(FreeNum) Then
 Line Input #FreeNum, fa '��ȡ��һ�е�����
 Exit Do '�˳�ѭ��
 End If
Loop
Close FreeNum

If fa <> "normal" Then
    End
End If
End Sub
