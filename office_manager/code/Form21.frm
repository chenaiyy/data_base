VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form21 
   Caption         =   "��������"
   ClientHeight    =   1950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4170
   LinkTopic       =   "Form21"
   ScaleHeight     =   1950
   ScaleWidth      =   4170
   StartUpPosition =   3  '����ȱʡ
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1560
      Top             =   1200
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      Connect         =   $"Form21.frx":0000
      OLEDBString     =   $"Form21.frx":0088
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
      Left            =   2760
      TabIndex        =   5
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ��"
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   1200
      Width           =   855
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      Left            =   2280
      TabIndex        =   3
      Top             =   720
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "���ϱ��"
      Height          =   375
      Index           =   1
      Left            =   720
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Caption         =   "��������"
      Height          =   375
      Index           =   0
      Left            =   720
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   2280
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "Form21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If Option1(1).Value = True Then
    Adodc1.RecordSource = "select ���ϱ��,�������� from ���������Ϣ�� where ���ϱ��='" & Combo1.Text & "'"
    Adodc1.Refresh
ElseIf Option1(0).Value = True Then
    Adodc1.RecordSource = "select ���ϱ��,�������� from ���������Ϣ�� where ��������='" & Combo2.Text & "'"
    Adodc1.Refresh
Else
    MsgBox "��û��ѡ�����ѡ��", , "��ʾ"
    GoTo Mark2
End If

If Adodc1.Recordset.EOF = True Then
    Dim name, number As String
    If Option1(1).Value = True Then
        Adodc1.RecordSource = "select ���ϱ��,�������� from ���ϱ� where ���ϱ��='" & Combo1.Text & "'"
        Adodc1.Refresh
    ElseIf Option1(0).Value = True Then
        Adodc1.RecordSource = "select ���ϱ��,�������� from ���ϱ� where ��������='" & Combo2.Text & "'"
        Adodc1.Refresh
    End If
    
    name = Adodc1.Recordset.Fields("��������").Value
    number = Adodc1.Recordset.Fields("���ϱ��").Value
    
    Adodc1.RecordSource = "�������ϱ�"
    Adodc1.Refresh
    
    Adodc1.Recordset.AddNew
    Adodc1.Recordset.Fields("���ϱ��").Value = number
    Adodc1.Recordset.Fields("��������").Value = name
    Adodc1.Recordset.Fields("������ְ����").Value = Form2.Label8.Caption
    Adodc1.Recordset.Fields("����ʱ��").Value = Form2.Label14.Caption
    Adodc1.Recordset.Update
    
    MsgBox "����ɹ�", , "��ʾ"
    
    Form21.Hide
    Form6.Show
Else
    MsgBox "�Բ�������������ѱ�����", , "��ʾ"
End If

Mark2:
End Sub

Private Sub Command2_Click()
Form21.Hide
Form2.Show
End Sub

Private Sub Form_Load()

Adodc1.RecordSource = "���ϱ�"
Adodc1.Refresh

Dim name, number As String
Adodc1.Recordset.MoveFirst
Do
name = Adodc1.Recordset.Fields("���ϱ��").Value
number = Adodc1.Recordset.Fields("��������").Value
Combo1.AddItem name
Combo2.AddItem number
Adodc1.Recordset.MoveNext
Loop While Adodc1.Recordset.EOF <> True

End Sub
