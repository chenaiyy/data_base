VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form2 
   Caption         =   "������Ϣ"
   ClientHeight    =   3930
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5235
   LinkTopic       =   "Form2"
   ScaleHeight     =   3930
   ScaleWidth      =   5235
   StartUpPosition =   3  '����ȱʡ
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   3000
      Top             =   0
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
      Connect         =   $"Form2.frx":0000
      OLEDBString     =   $"Form2.frx":0088
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
   Begin VB.CommandButton Command4 
      Caption         =   "�˳�"
      Height          =   375
      Left            =   1800
      TabIndex        =   22
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "��������"
      Height          =   375
      Left            =   240
      TabIndex        =   21
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Left            =   4560
      Top             =   240
   End
   Begin VB.CommandButton Command2 
      Caption         =   "����칫"
      Height          =   375
      Left            =   1800
      TabIndex        =   12
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�޸ĸ�������"
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   3000
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   720
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      Connect         =   $"Form2.frx":0110
      OLEDBString     =   $"Form2.frx":0198
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
   Begin VB.Frame Frame1 
      Caption         =   "������Ϣ"
      Height          =   2655
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2775
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Height          =   180
         Left            =   1080
         TabIndex        =   19
         Top             =   2160
         Width           =   90
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "���¹��ʣ�"
         Height          =   180
         Left            =   120
         TabIndex        =   18
         Top             =   2160
         Width           =   900
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Height          =   180
         Left            =   1200
         TabIndex        =   10
         Top             =   1800
         Width           =   90
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Height          =   180
         Left            =   1080
         TabIndex        =   9
         Top             =   1440
         Width           =   90
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Height          =   180
         Left            =   960
         TabIndex        =   8
         Top             =   1080
         Width           =   90
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Height          =   180
         Left            =   1080
         TabIndex        =   7
         Top             =   720
         Width           =   90
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "�������ţ�"
         Height          =   180
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Width           =   900
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "�Ա�"
         Height          =   180
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "ְ��"
         Height          =   180
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "ְ���ţ�"
         Height          =   180
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Height          =   180
         Left            =   1200
         TabIndex        =   2
         Top             =   360
         Width           =   90
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "��ӭ����"
         Height          =   180
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   720
      End
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   855
      Left            =   3120
      TabIndex        =   23
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label Label19 
      Caption         =   "Label19"
      Height          =   705
      Left            =   3120
      TabIndex        =   20
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Label16"
      Height          =   180
      Left            =   3600
      TabIndex        =   17
      Top             =   3600
      Width           =   630
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "����ʱ�䣺"
      Height          =   180
      Left            =   3240
      TabIndex        =   16
      Top             =   3000
      Width           =   900
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Label14"
      Height          =   225
      Left            =   3600
      TabIndex        =   15
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label Label13 
      Caption         =   "��˾��̬��"
      Height          =   255
      Left            =   3120
      TabIndex        =   14
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label12 
      Height          =   735
      Left            =   3120
      TabIndex        =   13
      Top             =   600
      Width           =   1935
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form3.Show
Form2.Hide
End Sub

Private Sub Command2_Click()
Dim bumen As String
bumen = Label11.Caption

Form2.Hide

Select Case bumen
    Case "�칫��"
        Form4.Show
    Case "���д�"
        Form5.Show
    Case "��Ϣ����"
        Form6.Show
    Case "�ۺϴ�"
        Form7.Show
End Select
End Sub

Private Sub Command3_Click()
Adodc2.RecordSource = "select ����ְ��ְ���� from ���������Ϣ�� where ����ְ��ְ����='" & Label8.Caption & "'"
Adodc2.Refresh

If Adodc2.Recordset.EOF = True Then
    Form2.Hide
    Form21.Show
Else
    MsgBox "�Բ������ϴν��ĵ����ϻ�û�й黹������ʱ���ܽ����κ�����", , "��ʾ"
End If
End Sub

Private Sub Command4_Click()
End
End Sub

Private Sub Form_Load()
Timer1.Interval = 300
 Dim a As Integer
    a = Weekday(Now(), vbMonday)
    Select Case a
          Case 1
            Label16.Caption = "����һ"
          Case 2
            Label16.Caption = "���ڶ�"
          Case 3
            Label16.Caption = "������"
          Case 4
            Label16.Caption = "������"
          Case 5
            Label16.Caption = "������"
          Case 6
            Label16.Caption = "������"
          Case 7
            Label16.Caption = "������"
    End Select
Dim number As String
number = Form1.Text1
Adodc1.RecordSource = "select ְ����,����,�Ա�,ְ��,���ڲ���,���� from ������Ϣ�� where ְ����='" & number & "'"
Adodc1.Refresh

Label2.Caption = Adodc1.Recordset.Fields("����").Value
Label8.Caption = number
Label9.Caption = Adodc1.Recordset.Fields("ְ��").Value
Label10.Caption = Adodc1.Recordset.Fields("�Ա�").Value
Label11.Caption = Adodc1.Recordset.Fields("���ڲ���").Value
Dim wage As Double
wage = Adodc1.Recordset.Fields("����").Value
Label18.Caption = Format(wage, "��0.00")

Adodc1.RecordSource = "��̬��"
Adodc1.Refresh
Adodc1.Recordset.MoveLast

Label12.Caption = Adodc1.Recordset.Fields("����").Value

Adodc2.RecordSource = "select �������� from ���������Ϣ�� where ����ְ��ְ����='" & number & "'"
Adodc2.Refresh

If Adodc2.Recordset.EOF = False Then
    Label19.Caption = Adodc2.Recordset.Fields("��������").Value
    Label19.Caption = "�����ĵ�" + Label19.Caption + "�Ѿ������׼���뼰ʱ����Ϣ������ȡ������"
    Label19.Visible = True
Else
    Label19.Visible = False
End If

Adodc2.RecordSource = "select ��Ŀ���� from �������Ŀ�� where ��Ŀ������ְ����='" & number & "'"
Adodc2.Refresh

If Adodc2.Recordset.EOF = False Then
    Label7.Caption = Adodc2.Recordset.Fields("��Ŀ����").Value
    Label7.Caption = "���ύ����Ŀ��" + Label7.Caption + "���Ѿ������׼���뼰ʱ��������ȡ�����Ŀ�ʽ�"
    Label7.Visible = True
Else
    Label7.Visible = False
End If

End Sub

Private Sub Timer1_Timer()
Label14.Caption = Now()
End Sub
