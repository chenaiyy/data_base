VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form16 
   Caption         =   "��Ŀ�������"
   ClientHeight    =   4545
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8145
   LinkTopic       =   "Form16"
   ScaleHeight     =   4545
   ScaleWidth      =   8145
   StartUpPosition =   3  '����ȱʡ
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   3000
      Top             =   0
      Visible         =   0   'False
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
      Connect         =   $"Form16.frx":0000
      OLEDBString     =   $"Form16.frx":0088
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
      Height          =   330
      Left            =   1800
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
      Connect         =   $"Form16.frx":0110
      OLEDBString     =   $"Form16.frx":019C
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
   Begin VB.CommandButton Command4 
      Caption         =   "����"
      Height          =   375
      Left            =   6120
      TabIndex        =   17
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "��һ��"
      Height          =   375
      Left            =   4080
      TabIndex        =   16
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "��˲�ͨ��"
      Height          =   375
      Left            =   2280
      TabIndex        =   15
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "���ͨ��"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   3960
      Width           =   1695
   End
   Begin VB.TextBox Text7 
      Height          =   1695
      Left            =   120
      TabIndex        =   13
      Text            =   "Text7"
      Top             =   2160
      Width           =   7335
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   5520
      TabIndex        =   12
      Text            =   "Text6"
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   2280
      TabIndex        =   11
      Text            =   "Text5"
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   5520
      TabIndex        =   10
      Text            =   "Text4"
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   5520
      TabIndex        =   9
      Text            =   "Text3"
      Top             =   240
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Text            =   "Text2"
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1440
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   240
      Width           =   1815
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
      Left            =   4080
      TabIndex        =   6
      Top             =   1440
      Width           =   1095
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
      Left            =   120
      TabIndex        =   5
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
      TabIndex        =   4
      Top             =   1920
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
      Left            =   4080
      TabIndex        =   3
      Top             =   840
      Width           =   1215
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
      Left            =   4080
      TabIndex        =   2
      Top             =   240
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
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Adodc1.RecordSource = "select * from �ύ��Ŀ�� where ��Ŀ��='" & Text1.Text & "'"
Adodc1.Refresh
Adodc1.Recordset.Delete

Adodc2.RecordSource = "�������Ŀ��"
Adodc2.Refresh

Adodc2.Recordset.AddNew
Adodc2.Recordset.Fields("��Ŀ��").Value = Text1.Text
Adodc2.Recordset.Fields("��Ŀ����").Value = Text2.Text
Adodc2.Recordset.Fields("��Ŀ�����ص�").Value = Text3.Text
Adodc2.Recordset.Fields("��Ŀ������ְ����").Value = Text5.Text
Adodc2.Recordset.Fields("��Ŀ����").Value = Format(Text6.Text)
Adodc2.Recordset.Fields("��Ŀ�ܽ�").Value = Text7.Text
Adodc2.Recordset.Fields("���ʱ��").Value = Text4.Text

Adodc2.Recordset.Update
MsgBox "�������", , "��ʾ"

End Sub

Private Sub Command2_Click()
Adodc1.RecordSource = "select * from �ύ��Ŀ�� where ��Ŀ��='" & Text1.Text & "'"
Adodc1.Refresh
Adodc1.Recordset.Delete

Adodc1.RecordSource = "���ڽ�����Ŀ��"
Adodc1.Refresh
Adodc1.Recordset.AddNew
Adodc1.Recordset.Fields("��Ŀ��").Value = Text1.Text
Adodc1.Recordset.Fields("��Ŀ����").Value = Text2.Text
Adodc1.Recordset.Fields("�����ص�").Value = Text3.Text
Adodc1.Recordset.Fields("��Ŀ������ְ����").Value = Text5.Text
Adodc1.Recordset.Fields("��Ŀ����").Value = Format(Text6.Text)
Adodc1.Recordset.Update

MsgBox "�������", , "��ʾ"

End Sub

Private Sub Command3_Click()
Adodc1.RecordSource = "�ύ��Ŀ��"
Adodc1.Refresh

If Adodc1.Recordset.EOF = True Then
    MsgBox "û����", , "��ʾ"
    Command1.Enabled = False
    Command2.Enabled = False
    Command3.Enabled = False
    GoTo Mark1
End If

Text1.Text = Adodc1.Recordset.Fields("��Ŀ��").Value
Text2.Text = Adodc1.Recordset.Fields("��Ŀ����").Value
Text6.Text = Format(Adodc1.Recordset.Fields("��Ŀ����").Value)
Text5.Text = Adodc1.Recordset.Fields("��Ŀ������ְ����").Value
Text3.Text = Adodc1.Recordset.Fields("��Ŀ�����ص�").Value
Text4.Text = Adodc1.Recordset.Fields("���ʱ��").Value
Text7.Text = Adodc1.Recordset.Fields("��Ŀ�ܽ�").Value
Mark1:
End Sub

Private Sub Command4_Click()
Form16.Hide
Form7.Show
End Sub

Private Sub Form_Load()
Adodc1.RecordSource = "�ύ��Ŀ��"
Adodc1.Refresh

If Adodc1.Recordset.EOF = True Then
    MsgBox "û����Ҫ��˵���Ŀ", , "��ʾ"
    Command1.Enabled = False
    Command2.Enabled = False
    Command3.Enabled = False
    GoTo Mark
End If

Text1.Text = Adodc1.Recordset.Fields("��Ŀ��").Value
Text2.Text = Adodc1.Recordset.Fields("��Ŀ����").Value
Text6.Text = Format(Adodc1.Recordset.Fields("��Ŀ����").Value)
Text5.Text = Adodc1.Recordset.Fields("��Ŀ������ְ����").Value
Text3.Text = Adodc1.Recordset.Fields("��Ŀ�����ص�").Value
Text4.Text = Adodc1.Recordset.Fields("���ʱ��").Value
Text7.Text = Adodc1.Recordset.Fields("��Ŀ�ܽ�").Value
Mark:

End Sub

