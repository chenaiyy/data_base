VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "�ۺϴ��칫����"
   ClientHeight    =   2760
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5850
   LinkTopic       =   "Form7"
   ScaleHeight     =   2760
   ScaleWidth      =   5850
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command4 
      Caption         =   "�˳�"
      Height          =   735
      Left            =   3240
      TabIndex        =   3
      Top             =   1680
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "��Ա����"
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "��Ŀ�������"
      Height          =   735
      Left            =   3240
      TabIndex        =   1
      Top             =   360
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��˾��̬����"
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form25.Show
Form7.Hide
End Sub

Private Sub Command2_Click()
Form16.Show
Form7.Hide
End Sub

Private Sub Command3_Click()
Form7.Hide
Form26.Show
End Sub

Private Sub Command4_Click()
End
End Sub

Private Sub Form_Load()

Dim zhiwu As String
zhiwu = Form2.Label9.Caption

If zhiwu <> "����" Then
    Command2.Enabled = False
    Command3.Enabled = False
End If

End Sub
