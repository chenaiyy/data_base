VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form11 
   Caption         =   "��Ŀ�������"
   ClientHeight    =   6165
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6900
   LinkTopic       =   "Form11"
   ScaleHeight     =   6165
   ScaleWidth      =   6900
   StartUpPosition =   3  '����ȱʡ
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   2640
      Top             =   1560
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"Form11.frx":0000
      OLEDBString     =   $"Form11.frx":008C
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ȡ��"
      Height          =   495
      Left            =   3960
      TabIndex        =   11
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ��"
      Height          =   495
      Left            =   840
      TabIndex        =   10
      Top             =   5640
      Width           =   1335
   End
   Begin VB.TextBox Text6 
      Height          =   1575
      Left            =   360
      TabIndex        =   9
      Top             =   3960
      Width           =   6375
   End
   Begin VB.TextBox Text5 
      Height          =   1095
      Left            =   360
      TabIndex        =   8
      Top             =   2160
      Width           =   6375
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   4800
      TabIndex        =   7
      Top             =   240
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "��Ŀ���걨����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "��Ŀ�ľ�������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "�����ص�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "��Ŀ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "��Ŀ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If Text1.Text = "" Then
    MsgBox "��������Ŀ����", , "��ʾ"
    GoTo Label1
End If

If Text2.Text = "" Then
    MsgBox "��������Ŀ����", , "��ʾ"
    GoTo Label1
End If

If Text3.Text = "" Then
    MsgBox "��������Ŀ�����ص�", , "��ʾ"
    GoTo Label1
End If

If Text5.Text = "" Then
    MsgBox "��������Ŀ��������", , "��ʾ"
    GoTo Label1
End If

If Text6.Text = "" Then
    MsgBox "����д��Ŀ���걨����", , "��ʾ"
    GoTo Label1
End If

Adodc1.RecordSource = "��Ŀ�����"
Adodc1.Refresh
Adodc1.Recordset.AddNew

Adodc1.Recordset.Fields("��Ŀ����") = Text1.Text
Adodc1.Recordset.Fields("��Ŀ����") = Format(Text2.Text)
Adodc1.Recordset.Fields("�����ص�") = Text3.Text
Adodc1.Recordset.Fields("��Ŀ������ְ����") = Form2.Label8.Caption
Adodc1.Recordset.Fields("��Ŀ��������") = Text5.Text
Adodc1.Recordset.Fields("��Ŀ�걨����") = Text6.Text
Adodc1.Recordset.Update

MsgBox "��Ŀ�걨���", , "��ʾ"
Form11.Hide
Form5.Show

Label1:

End Sub

Private Sub Command2_Click()
Form11.Hide
Form5.Show
End Sub
