VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form26 
   Caption         =   "��Ա����"
   ClientHeight    =   3165
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6420
   LinkTopic       =   "Form26"
   ScaleHeight     =   3165
   ScaleWidth      =   6420
   StartUpPosition =   3  '����ȱʡ
   Begin VB.ComboBox Combo2 
      Height          =   300
      ItemData        =   "Form26.frx":0000
      Left            =   4320
      List            =   "Form26.frx":000D
      TabIndex        =   13
      Top             =   1800
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "Form26.frx":0027
      Left            =   4320
      List            =   "Form26.frx":0037
      Sorted          =   -1  'True
      TabIndex        =   12
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ȡ��"
      Height          =   495
      Left            =   4200
      TabIndex        =   11
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ��"
      Height          =   495
      Left            =   360
      TabIndex        =   10
      Top             =   2400
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   2400
      Top             =   2520
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
      Connect         =   $"Form26.frx":005D
      OLEDBString     =   $"Form26.frx":00E5
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
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1320
      TabIndex        =   9
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   1320
      TabIndex        =   8
      Top             =   1800
      Width           =   90
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   1320
      TabIndex        =   7
      Top             =   1080
      Width           =   90
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   4320
      TabIndex        =   6
      Top             =   360
      Width           =   90
   End
   Begin VB.Label Label1 
      Caption         =   "��ְ��"
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
      Index           =   5
      Left            =   3240
      TabIndex        =   5
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "�ֲ���"
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
      Left            =   3240
      TabIndex        =   4
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "ԭ����"
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
      Index           =   3
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "ԭְ��"
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
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "����"
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
      Index           =   1
      Left            =   3240
      TabIndex        =   1
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "ְ����"
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
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "Form26"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Combo1.Text = "�칫��" And Combo2.Text = "����" Then
    MsgBox "�칫�Ҳ����ڴ���ְ��", , "��ʾ"
    GoTo Mark1
End If

If Combo1.Text = "���д�" And Combo2.Text = "����" Then
    MsgBox "���д�����������ְ��", , "��ʾ"
    GoTo Mark1
End If

If Combo1.Text = "�ۺϴ�" And Combo2.Text = "����" Then
    MsgBox "�ۺϴ�����������ְ��", , "��ʾ"
    GoTo Mark1
End If

If Combo1.Text = "��Ϣ����" And Combo2.Text = "����" Then
    MsgBox "��Ϣ���Ĳ����ڴ���ְ��", , "��ʾ"
    GoTo Mark1
End If

Adodc1.Recordset.Fields("���ڲ���").Value = Combo1.Text
Adodc1.Recordset.Fields("ְ��").Value = Combo2.Text
Adodc1.Recordset.Update

MsgBox "�������", , "��ʾ"
Form26.Hide
Form7.Show

Mark1:
End Sub

Private Sub Command2_Click()
Form26.Hide
Form7.Show
End Sub

Private Sub Form_Load()
Combo1.Enabled = False
Combo2.Enabled = False
End Sub

Private Sub Text1_LostFocus()
If Text1.Text = "" Then
    MsgBox "�����뱻������Ա��ְ����", , "��ʾ"
Else
    Adodc1.RecordSource = "select ְ����,����,ְ��,���ڲ��� from ������Ϣ�� where ְ����='" & Text1.Text & "'"
    Adodc1.Refresh
    If Adodc1.Recordset.EOF = True Then
        MsgBox "�������ְ���Ų�����", , "��ʾ"
        Text1.Text = ""
        GoTo Mark
    Else
        Label3.Caption = Adodc1.Recordset.Fields("����").Value
        Label4.Caption = Adodc1.Recordset.Fields("���ڲ���").Value
        Label5.Caption = Adodc1.Recordset.Fields("ְ��").Value
        Combo1.Enabled = True
        Combo2.Enabled = True
    End If
End If
Mark:
End Sub
