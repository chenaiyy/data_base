VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form24 
   Caption         =   "���Ϲ黹"
   ClientHeight    =   1545
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4110
   LinkTopic       =   "Form24"
   ScaleHeight     =   1545
   ScaleWidth      =   4110
   StartUpPosition =   3  '����ȱʡ
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   1200
      Top             =   960
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
      Connect         =   $"Form24.frx":0000
      OLEDBString     =   $"Form24.frx":0088
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
      Left            =   2400
      TabIndex        =   3
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ��"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "���ϱ��"
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
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form24"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text <> "" Then
    Adodc1.RecordSource = "select * from ���������Ϣ�� where ���ϱ��='" & Text1.Text & "'"
    Adodc1.Refresh
    If Adodc1.Recordset.EOF = False Then
        Adodc1.Recordset.Delete
        MsgBox "�����ɹ�", , "��ʾ"
        i = MsgBox("�Ƿ��������", vbYesNo, "��ʾ")
        If i = vbNo Then
            Form24.Hide
            Form6.Show
        Else
            Text1.Text = ""
        End If
    Else
        MsgBox "������û�����������Ͽ���û�и����ϣ���˶�", , "��ʾ"
    End If
Else
    MsgBox "��û���������ϱ��", , "��ʾ"
End If

End Sub

Private Sub Command2_Click()
Form24.Hide
Form6.Show
End Sub
