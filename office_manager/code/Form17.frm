VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form17 
   Caption         =   "��˲�ͨ��"
   ClientHeight    =   3585
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5625
   LinkTopic       =   "Form17"
   ScaleHeight     =   3585
   ScaleWidth      =   5625
   StartUpPosition =   3  '����ȱʡ
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   2280
      Top             =   3000
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
      Connect         =   $"Form17.frx":0000
      OLEDBString     =   $"Form17.frx":008C
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
      Left            =   600
      Top             =   3000
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
      Connect         =   $"Form17.frx":0118
      OLEDBString     =   $"Form17.frx":01A4
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
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ��"
      Height          =   495
      Left            =   4080
      TabIndex        =   2
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   2415
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   5175
   End
   Begin VB.Label Label1 
      Caption         =   "��˲���ͨ����ԭ��"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "Form17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "" Then
    MsgBox "û�н���ԭ������д", , "��ʾ"
    GoTo Label3
Else
    Adodc1.RecordSource = "��Ŀ��˲���˵����"
    Adodc1.Refresh
    
    Adodc1.Recordset.AddNew
    Adodc1.Recordset.Fields("ְ����").Value = Form13.Text4
    Adodc1.Recordset.Fields("��Ŀ����").Value = Form13.Text1
    Adodc1.Recordset.Fields("��ͨ��ԭ��").Value = Text1.Text
    Adodc1.Recordset.Update
    
    Adodc2.RecordSource = "select * from ��Ŀ����� where ��Ŀ������ְ����='" & Form13.Text4 & "'"
    Adodc2.Refresh
    Adodc2.Recordset.Delete
    
    MsgBox "�������", , "��ʾ"
    Form17.Hide
    Form13.Show
End If
Label3:
End Sub
