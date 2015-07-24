VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form19 
   Caption         =   "资料申请信息"
   ClientHeight    =   2715
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7395
   LinkTopic       =   "Form19"
   ScaleHeight     =   2715
   ScaleWidth      =   7395
   StartUpPosition =   3  '窗口缺省
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   840
      Top             =   1560
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
      Connect         =   $"Form19.frx":0000
      OLEDBString     =   $"Form19.frx":0088
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
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
      Caption         =   "返回"
      Height          =   615
      Left            =   6240
      TabIndex        =   11
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "申请不通过"
      Height          =   615
      Left            =   2280
      TabIndex        =   10
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "申请通过"
      Height          =   615
      Left            =   480
      TabIndex        =   9
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "下一个"
      Height          =   615
      Left            =   4320
      TabIndex        =   8
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   240
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "申请资料编号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   3600
      TabIndex        =   3
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "资料名称"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   3720
      TabIndex        =   2
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "申请时间"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   600
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "申请人职工号"
      BeginProperty Font 
         Name            =   "宋体"
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
      Width           =   1575
   End
End
Attribute VB_Name = "Form19"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Adodc1.RecordSource = "申请资料表"
Adodc1.Refresh

If Adodc1.Recordset.EOF = True Then
    MsgBox "没有下一条记录", , "提示"
    Command1.Enabled = False
    Command2.Enabled = False
    Command3.Enabled = False
Else
    Adodc1.Recordset.MoveNext
End If
End Sub

Private Sub Command2_Click()

Adodc1.RecordSource = "资料外借信息表"
Adodc1.Refresh
Adodc1.Recordset.AddNew
Adodc1.Recordset.Fields("资料编号").Value = Text3.Text
Adodc1.Recordset.Fields("借阅职工职工号").Value = Text1.Text
Adodc1.Recordset.Fields("借出时间").Value = Text2.Text
Adodc1.Recordset.Fields("资料名称").Value = Text4.Text
Adodc1.Recordset.Update

Adodc1.RecordSource = "select * from 申请资料表 where 资料编号='" & Text3.Text & "'"
Adodc1.Refresh
Adodc1.Recordset.Delete

MsgBox "操作完成", , "提示"

End Sub

Private Sub Command3_Click()

Adodc1.RecordSource = "select * from 申请资料表 where 资料编号='" & Text3.Text & "'"
Adodc1.Refresh
Adodc1.Recordset.Delete

MsgBox "操作完成", , "提示"
End Sub

Private Sub Command5_Click()
Form19.Hide
Form6.Show
End Sub

Private Sub Form_Load()
Adodc1.RecordSource = "申请资料表"
Adodc1.Refresh

If Adodc1.Recordset.EOF = True Then
    Command1.Enabled = False
    Command2.Enabled = False
    Command3.Enabled = False
    MsgBox "没有人申请资料", , "提示"
Else
    Text1.Text = Adodc1.Recordset.Fields("申请人职工号").Value
    Text2.Text = Adodc1.Recordset.Fields("申请时间").Value
    Text3.Text = Adodc1.Recordset.Fields("资料编号").Value
    Text4.Text = Adodc1.Recordset.Fields("资料名称").Value
End If
End Sub
