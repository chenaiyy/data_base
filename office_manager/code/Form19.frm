VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form19 
   Caption         =   "����������Ϣ"
   ClientHeight    =   2715
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7395
   LinkTopic       =   "Form19"
   ScaleHeight     =   2715
   ScaleWidth      =   7395
   StartUpPosition =   3  '����ȱʡ
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   840
      Top             =   1560
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
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
      Connect         =   $"Form19.frx":0000
      OLEDBString     =   $"Form19.frx":0088
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
   Begin VB.CommandButton Command5 
      Caption         =   "����"
      Height          =   615
      Left            =   6240
      TabIndex        =   11
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "���벻ͨ��"
      Height          =   615
      Left            =   2280
      TabIndex        =   10
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "����ͨ��"
      Height          =   615
      Left            =   480
      TabIndex        =   9
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��һ��"
      Height          =   615
      Left            =   4320
      TabIndex        =   8
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   240
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "�������ϱ��"
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
      Index           =   3
      Left            =   3600
      TabIndex        =   3
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "��������"
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
      Index           =   2
      Left            =   3720
      TabIndex        =   2
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "����ʱ��"
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
      Index           =   1
      Left            =   600
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "������ְ����"
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
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "Form19"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Adodc1.RecordSource = "�������ϱ�"
Adodc1.Refresh

If Adodc1.Recordset.EOF = True Then
    MsgBox "û����һ����¼", , "��ʾ"
    Command1.Enabled = False
    Command2.Enabled = False
    Command3.Enabled = False
Else
    Adodc1.Recordset.MoveNext
End If
End Sub

Private Sub Command2_Click()

Adodc1.RecordSource = "���������Ϣ��"
Adodc1.Refresh
Adodc1.Recordset.AddNew
Adodc1.Recordset.Fields("���ϱ��").Value = Text3.Text
Adodc1.Recordset.Fields("����ְ��ְ����").Value = Text1.Text
Adodc1.Recordset.Fields("���ʱ��").Value = Text2.Text
Adodc1.Recordset.Fields("��������").Value = Text4.Text
Adodc1.Recordset.Update

Adodc1.RecordSource = "select * from �������ϱ� where ���ϱ��='" & Text3.Text & "'"
Adodc1.Refresh
Adodc1.Recordset.Delete

MsgBox "�������", , "��ʾ"

End Sub

Private Sub Command3_Click()

Adodc1.RecordSource = "select * from �������ϱ� where ���ϱ��='" & Text3.Text & "'"
Adodc1.Refresh
Adodc1.Recordset.Delete

MsgBox "�������", , "��ʾ"
End Sub

Private Sub Command5_Click()
Form19.Hide
Form6.Show
End Sub

Private Sub Form_Load()
Adodc1.RecordSource = "�������ϱ�"
Adodc1.Refresh

If Adodc1.Recordset.EOF = True Then
    Command1.Enabled = False
    Command2.Enabled = False
    Command3.Enabled = False
    MsgBox "û������������", , "��ʾ"
Else
    Text1.Text = Adodc1.Recordset.Fields("������ְ����").Value
    Text2.Text = Adodc1.Recordset.Fields("����ʱ��").Value
    Text3.Text = Adodc1.Recordset.Fields("���ϱ��").Value
    Text4.Text = Adodc1.Recordset.Fields("��������").Value
End If
End Sub
