VERSION 5.00
Begin VB.Form place 
   Caption         =   "位置调节"
   ClientHeight    =   3030
   ClientLeft      =   17970
   ClientTop       =   4500
   ClientWidth     =   3390
   LinkTopic       =   "Form3"
   ScaleHeight     =   3030
   ScaleWidth      =   3390
   Begin VB.CommandButton Command5 
      Caption         =   "确认"
      Height          =   855
      Left            =   1200
      TabIndex        =   4
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "↓"
      Height          =   855
      Left            =   1200
      TabIndex        =   3
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "→"
      Height          =   855
      Left            =   2280
      TabIndex        =   2
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "←"
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "↑"
      Height          =   855
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "place"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.top = Form1.top - 200
End Sub

Private Sub Command2_Click()
Form1.left = Form1.left - 200
End Sub

Private Sub Command3_Click()
Form1.left = Form1.left + 200
End Sub

Private Sub Command4_Click()
Form1.top = Form1.top + 200
End Sub

Private Sub Command5_Click()
Unload place
End Sub
