VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form10 
   Caption         =   "���ʸ��Ľ���"
   ClientHeight    =   2220
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3180
   LinkTopic       =   "Form10"
   ScaleHeight     =   2220
   ScaleWidth      =   3180
   StartUpPosition =   3  '����ȱʡ
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   960
      Top             =   1320
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
      Connect         =   $"Form10.frx":0000
      OLEDBString     =   $"Form10.frx":0088
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
      Caption         =   "����"
      Height          =   495
      Left            =   1920
      TabIndex        =   5
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ��"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "���ʱ䶯��"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "ԭ�й��ʣ�"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim i As Integer
If Text1.Text <> "" Then
    i = MsgBox("�Ƿ�ȷ�����Ĺ���", vbOKCancel, "��ʾ")
    If i = 1 Then
        Adodc1.Recordset.Fields("����").Value = Format(Text1.Text)
        Adodc1.Refresh
        Label3.Caption = Text1.Text
    End If
Else
    MsgBox "������Ҫ���ĵ�ֵ", , "��ʾ"
End If
If Text1.Text <> "" Then
Text1.Text = ""
End If
End Sub

Private Sub Command2_Click()
Form10.Hide
Form4.Show
End Sub

Private Sub Form_Load()
Adodc1.RecordSource = "select ְ����,���� from ������Ϣ�� where ְ����='" & Form4.Text2 & "'"
Adodc1.Refresh

Dim wage As Double
wage = Adodc1.Recordset.Fields("����").Value
Label3.Caption = Format(wage, "��0.00")
End Sub
