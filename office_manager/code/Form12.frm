VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form12 
   Caption         =   "��Ŀ�ύ����"
   ClientHeight    =   5385
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7230
   LinkTopic       =   "Form12"
   ScaleHeight     =   5385
   ScaleWidth      =   7230
   StartUpPosition =   3  '����ȱʡ
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   375
      Left            =   3840
      Top             =   4560
      Width           =   1200
      _ExtentX        =   2117
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
      Connect         =   $"Form12.frx":0000
      OLEDBString     =   $"Form12.frx":0088
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc3"
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   2400
      Top             =   4920
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
      Connect         =   $"Form12.frx":0110
      OLEDBString     =   $"Form12.frx":019C
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc2"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   2400
      Top             =   4560
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
      Connect         =   $"Form12.frx":0228
      OLEDBString     =   $"Form12.frx":02B4
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
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   5880
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "Text7"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ȡ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   13
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�ύ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   12
      Top             =   4680
      Width           =   1575
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "Text6"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      Height          =   2535
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Width           =   6855
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   5520
      TabIndex        =   7
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "Text3"
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "Text2"
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   120
      Width           =   1455
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
      Index           =   6
      Left            =   4800
      TabIndex        =   14
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label1 
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
      Index           =   5
      Left            =   3960
      TabIndex        =   10
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "��Ŀ������ְ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   1440
      TabIndex        =   8
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "��Ŀ�ܽ�"
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
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "���ʱ��"
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
      Left            =   3960
      TabIndex        =   2
      Top             =   840
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
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "��Ŀ���"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text4.Text = "" Then
    MsgBox "���������ʱ��", , "��ʾ"
    GoTo Label4
End If

If Text5.Text = "" Then
    MsgBox "��û�������ܽᱨ��", , "��ʾ"
    GoTo Label4
End If

Adodc2.RecordSource = "�ύ��Ŀ��"
Adodc2.Refresh
Adodc2.Recordset.AddNew

Adodc2.Recordset.Fields("��Ŀ��").Value = Text1.Text
Adodc2.Recordset.Fields("��Ŀ�����ص�").Value = Text3.Text
Adodc2.Recordset.Fields("���ʱ��").Value = Text4.Text
Adodc2.Recordset.Fields("��Ŀ����").Value = Text2.Text
Adodc2.Recordset.Fields("��Ŀ������ְ����").Value = Text6.Text
Adodc2.Recordset.Fields("��Ŀ�ܽ�").Value = Text5.Text
Adodc2.Recordset.Fields("��Ŀ����").Value = Format(Text7.Text)
Adodc2.Recordset.Update

Adodc1.Recordset.Delete

Adodc3.RecordSource = "��Ŀ��ͼ"
Adodc3.Refresh
Adodc3.Recordset.Fields("��ǰ������Ŀ").Value = "��"
Adodc3.Recordset.Update

MsgBox "�������", , "��ʾ"
Form12.Hide
Form5.Show
Label4:
End Sub

Private Sub Command2_Click()
Form12.Hide
Form5.Show
End Sub

Private Sub Form_Load()
Adodc1.RecordSource = "select * from ���ڽ�����Ŀ�� where ��Ŀ������ְ����='" & Form2.Label8.Caption & "'"
Adodc1.Refresh
If Adodc1.Recordset.EOF = True Then
    MsgBox "��û����Ŀ�����ύ", , "��ʾ"
    GoTo Mark
End If

Text1.Text = Adodc1.Recordset.Fields("��Ŀ��").Value
Text2.Text = Adodc1.Recordset.Fields("��Ŀ����").Value
Text7.Text = Format(Adodc1.Recordset.Fields("��Ŀ����").Value)
Text6.Text = Adodc1.Recordset.Fields("��Ŀ������ְ����").Value
Text3.Text = Adodc1.Recordset.Fields("�����ص�").Value

Mark:

End Sub
