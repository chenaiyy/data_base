VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "��Ϣ���İ칫����"
   ClientHeight    =   3630
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5385
   LinkTopic       =   "Form6"
   ScaleHeight     =   3630
   ScaleWidth      =   5385
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command5 
      Caption         =   "���Ϲ黹"
      Height          =   855
      Left            =   2880
      TabIndex        =   4
      Top             =   1680
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "�鿴�������ϼ�¼"
      Height          =   855
      Left            =   360
      TabIndex        =   3
      Top             =   1680
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "�˳�"
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   2760
      Width           =   4695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�������"
      Height          =   855
      Left            =   2880
      TabIndex        =   1
      Top             =   360
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�鿴����"
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form19.Show
Form6.Hide
End Sub

Private Sub Command2_Click()
Form20.Show
Form6.Hide
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Command4_Click()
Form23.Show
Form6.Hide
End Sub

Private Sub Command5_Click()
Form6.Hide
Form24.Show
End Sub

Private Sub Form_Load()
Dim zhiwu As String
zhiwu = Form2.Label9.Caption

If zhiwu <> "����" Then
    Command2.Enabled = False
End If
End Sub
