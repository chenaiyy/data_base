VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form3 
   Caption         =   "�������"
   ClientHeight    =   2925
   ClientLeft      =   1125
   ClientTop       =   450
   ClientWidth     =   3285
   LinkTopic       =   "Form3"
   ScaleHeight     =   2925
   ScaleWidth      =   3285
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command2 
      Caption         =   "ȡ��"
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ��"
      Height          =   300
      Left            =   240
      TabIndex        =   6
      Top             =   1920
      Width           =   855
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   480
      Top             =   2280
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      Connect         =   $"Form3.frx":0000
      OLEDBString     =   $"Form3.frx":0088
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
   Begin VB.TextBox Text3 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "ȷ��������"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "����������"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "����ԭ����"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim namenumber As String
namenumber = Form2.Label8

Adodc1.RecordSource = "select ְ����,���� from ������Ϣ�� where ְ����='" & namenumber & "'"
Adodc1.Refresh

If Text1.Text <> Adodc1.Recordset.Fields("����").Value Then
    MsgBox "�����ԭ�������", , "��ʾ"
    GoTo Label
End If

If Text2.Text = "" Then
    MsgBox "���벻��Ϊ��", , "��ʾ"
    GoTo Label
End If

If Text3.Text = "" Then
    MsgBox "��ȷ������", , "��ʾ"
    GoTo Label
End If

If Text2.Text = Text3.Text Then
    Adodc1.Recordset.Fields("����").Value = Text2.Text
    Adodc1.Recordset.Update
    MsgBox "�����Ѿ��޸ģ�������ס����������", , "��ʾ"
    Form3.Hide
    Form2.Show
Else
    MsgBox "����ȷ�ϴ���", , "��ʾ"
End If
Label:
End Sub

Private Sub Command2_Click()
Form3.Hide
Form2.Show
End Sub
