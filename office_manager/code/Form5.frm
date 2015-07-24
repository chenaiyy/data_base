VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "科研处办公界面"
   ClientHeight    =   5745
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3735
   LinkTopic       =   "Form5"
   ScaleHeight     =   5745
   ScaleWidth      =   3735
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command6 
      Caption         =   "审核未通过浏览"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   5
      Top             =   3120
      Width           =   2775
   End
   Begin VB.CommandButton Command5 
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   4
      Top             =   5040
      Width           =   2775
   End
   Begin VB.CommandButton Command4 
      Caption         =   "审核通过项目浏览"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   3
      Top             =   2160
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "项目审核"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   2
      Top             =   4080
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "项目提交"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   1
      Top             =   1200
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "项目申请"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form5.Hide
Form11.Show
End Sub

Private Sub Command2_Click()
Form5.Hide
Form12.Show
End Sub

Private Sub Command3_Click()
Form5.Hide
Form13.Show
End Sub

Private Sub Command4_Click()
Form5.Hide
Form14.Show
End Sub

Private Sub Command5_Click()
End
End Sub

Private Sub Command6_Click()
Form5.Hide
Form18.Show
End Sub

Private Sub Form_Load()
Dim zhiwu As String
zhiwu = Form2.Label9.Caption

If zhiwu <> "处长" Then
    Command3.Enabled = False
Else
    Command1.Enabled = False
    Command2.Enabled = False
    Command4.Enabled = False
    Command6.Enabled = False
End If
End Sub
