VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form25 
   Caption         =   "��̬��������"
   ClientHeight    =   3135
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5085
   LinkTopic       =   "Form25"
   ScaleHeight     =   3135
   ScaleWidth      =   5085
   StartUpPosition =   3  '����ȱʡ
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   3360
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
      Connect         =   $"Form25.frx":0000
      OLEDBString     =   $"Form25.frx":0088
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
      Left            =   3600
      TabIndex        =   3
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ��"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   2055
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   4815
   End
   Begin VB.Label Label1 
      Caption         =   "�ļ�����"
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
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form25"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If Text2.Text = "" Then
    MsgBox "��û����������", , "��ʾ"
    GoTo Mark
End If

Dim strSNum As String
strSNum = Format(Now, "yyyy-mm-dd-hhmmss")

Adodc1.RecordSource = "��̬��"
Adodc1.Refresh
Adodc1.Recordset.AddNew
Adodc1.Recordset.Fields("��̬���").Value = strSNum
Adodc1.Recordset.Fields("����").Value = Text2.Text
Adodc1.Recordset.Update

MsgBox "�����ɹ�", , "��ʾ"
Form25.Hide
Form7.Show
Mark:
End Sub

Private Sub Command2_Click()
Form25.Hide
Form7.Show
End Sub
