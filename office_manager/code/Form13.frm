VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form13 
   Caption         =   "��Ŀ��˽���"
   ClientHeight    =   5805
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6840
   LinkTopic       =   "Form13"
   ScaleHeight     =   5805
   ScaleWidth      =   6840
   StartUpPosition =   3  '����ȱʡ
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   4800
      Top             =   5520
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
      Connect         =   $"Form13.frx":0000
      OLEDBString     =   $"Form13.frx":008C
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
      Height          =   330
      Left            =   2400
      Top             =   5400
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
      Connect         =   $"Form13.frx":0118
      OLEDBString     =   $"Form13.frx":01A0
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
   Begin VB.CommandButton Command5 
      Caption         =   "������һ��"
      Height          =   495
      Left            =   5520
      TabIndex        =   16
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "��һ��"
      Height          =   495
      Left            =   4440
      TabIndex        =   15
      Top             =   4920
      Width           =   855
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   360
      Top             =   5400
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   $"Form13.frx":0228
      OLEDBString     =   $"Form13.frx":02B4
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "��Ŀ�����"
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
   Begin VB.CommandButton Command3 
      Caption         =   "��˲�ͨ��"
      Height          =   495
      Left            =   3120
      TabIndex        =   14
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "���ͨ��"
      Height          =   495
      Left            =   1800
      TabIndex        =   13
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�鿴��������Ϣ"
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   4920
      Width           =   1455
   End
   Begin VB.TextBox Text6 
      DataField       =   "��Ŀ�걨����"
      DataSource      =   "Adodc1"
      Height          =   1095
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "Text6"
      Top             =   3600
      Width           =   6375
   End
   Begin VB.TextBox Text5 
      DataField       =   "��Ŀ��������"
      DataSource      =   "Adodc1"
      Height          =   1095
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "Text5"
      Top             =   1800
      Width           =   6375
   End
   Begin VB.TextBox Text4 
      DataField       =   "��Ŀ������ְ����"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "Text4"
      Top             =   1320
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      DataField       =   "�����ص�"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "Text3"
      Top             =   240
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      DataField       =   "��Ŀ����"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      DataField       =   "��Ŀ����"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "��Ŀ�걨����"
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
      Left            =   120
      TabIndex        =   5
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "��Ŀ��������"
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
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   1815
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
      Height          =   375
      Index           =   3
      Left            =   3360
      TabIndex        =   3
      Top             =   840
      Width           =   2175
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
      Index           =   2
      Left            =   3360
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
      Left            =   240
      TabIndex        =   1
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
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form13.Hide
Form15.Show
End Sub

Private Sub Command2_Click()
Dim strSNum As String
strSNum = Format(Now, "yyyymmddhhmmss")

Adodc3.RecordSource = "���ڽ�����Ŀ��"
Adodc3.Refresh

Adodc3.Recordset.AddNew
Adodc3.Recordset.Fields("��Ŀ��").Value = strSNum
Adodc3.Recordset.Fields("��Ŀ����").Value = Text1.Text
Adodc3.Recordset.Fields("�����ص�").Value = Text3.Text
Adodc3.Recordset.Fields("��Ŀ������ְ����").Value = Text4.Text
Adodc3.Recordset.Fields("��Ŀ����").Value = Format(Text2.Text)
Adodc3.Recordset.Update

Adodc2.RecordSource = "select ְ����,��ǰ������Ŀ from ������Ϣ�� where ְ����='" & Text4 & "'"
Adodc2.Refresh
Adodc2.Recordset.Fields("��ǰ������Ŀ").Value = strSNum
Adodc2.Recordset.Update

Adodc3.RecordSource = "select * from ��Ŀ����� where ��Ŀ������ְ����='" & Text4 & "'"
Adodc3.Refresh
Adodc3.Recordset.Delete

MsgBox "�������", , "��ʾ"
Adodc1.Refresh
End Sub

Private Sub Command3_Click()
Form13.Hide
Form17.Show
End Sub

Private Sub Command4_Click()
If Adodc1.Recordset.EOF <> True Then
    Adodc1.Recordset.MoveNext
Else
    MsgBox "��û���걨��Ϣ", , "��ʾ"
    Command1.Enabled = False
    Command2.Enabled = False
    Command3.Enabled = False
    Command4.Enabled = False
End If
End Sub

Private Sub Command5_Click()
Form13.Hide
Form5.Show
End Sub

Private Sub Form_Load()
If Adodc1.Recordset.EOF = True Then
    MsgBox "û���걨��Ϣ", , "��ʾ"
    Command1.Enabled = False
    Command2.Enabled = False
    Command3.Enabled = False
    Command4.Enabled = False
End If
End Sub
