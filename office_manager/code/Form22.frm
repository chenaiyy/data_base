VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form22 
   Caption         =   "人员添加"
   ClientHeight    =   3885
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6375
   LinkTopic       =   "Form22"
   ScaleHeight     =   3885
   ScaleWidth      =   6375
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "重置"
      Height          =   420
      Left            =   4560
      TabIndex        =   20
      Top             =   3000
      Width           =   855
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2400
      Top             =   2520
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
      Connect         =   $"Form22.frx":0000
      OLEDBString     =   $"Form22.frx":0088
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
   Begin VB.CommandButton Command2 
      Caption         =   "确定"
      Height          =   375
      Left            =   4560
      TabIndex        =   18
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "取消"
      Height          =   375
      Left            =   5520
      TabIndex        =   17
      Top             =   3000
      Width           =   735
   End
   Begin VB.ComboBox Combo4 
      Height          =   300
      ItemData        =   "Form22.frx":0110
      Left            =   4800
      List            =   "Form22.frx":0120
      TabIndex        =   16
      Top             =   2040
      Width           =   1215
   End
   Begin VB.ComboBox Combo3 
      Height          =   300
      ItemData        =   "Form22.frx":0146
      Left            =   4800
      List            =   "Form22.frx":0153
      TabIndex        =   15
      Top             =   1440
      Width           =   1215
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      ItemData        =   "Form22.frx":016D
      Left            =   4800
      List            =   "Form22.frx":0180
      TabIndex        =   14
      Top             =   840
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "Form22.frx":01AA
      Left            =   4800
      List            =   "Form22.frx":01B4
      TabIndex        =   13
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   735
      Left            =   360
      TabIndex        =   12
      Top             =   3000
      Width           =   3975
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   1320
      TabIndex        =   11
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1320
      TabIndex        =   10
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1320
      TabIndex        =   9
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1320
      TabIndex        =   19
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "学历"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   3840
      TabIndex        =   8
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "职务"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   3840
      TabIndex        =   7
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "联系电话"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "住址"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   360
      TabIndex        =   5
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "所属部门"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   3600
      TabIndex        =   4
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "性别"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   3840
      TabIndex        =   3
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "年龄"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   2
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "职工号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "姓名"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "Form22"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form22.Hide
Form4.Show
End Sub

Private Sub Command2_Click()
If Combo1.Text = "办公室" And Combo2.Text = "处长" Then
    MsgBox "办公室不存在处长职务", , "提示"
    GoTo Mark1
End If

If Combo1.Text = "科研处" And Combo2.Text = "主任" Then
    MsgBox "科研处不存在主任职务", , "提示"
    GoTo Mark1
End If

If Combo1.Text = "综合处" And Combo2.Text = "主任" Then
    MsgBox "综合处不存在主任职务", , "提示"
    GoTo Mark1
End If

If Combo1.Text = "信息中心" And Combo2.Text = "处长" Then
    MsgBox "信息中心不存在处长职务", , "提示"
    GoTo Mark1
End If

Adodc1.Recordset.AddNew
Adodc1.Recordset.Fields("姓名").Value = Text1.Text
Adodc1.Recordset.Fields("年龄").Value = Text2.Text
Adodc1.Recordset.Fields("职工号").Value = Label2.Caption
Adodc1.Recordset.Fields("联系电话").Value = Text4.Text
Adodc1.Recordset.Fields("住址").Value = Text5.Text
Adodc1.Recordset.Fields("性别").Value = Combo1.Text
Adodc1.Recordset.Fields("学历").Value = Combo2.Text
Adodc1.Recordset.Fields("职务").Value = Combo3.Text
Adodc1.Recordset.Fields("所在部门").Value = Combo4.Text
Adodc1.Recordset.Update

MsgBox "添加信息成功", , "提示"
Form22.Hide
Form4.Show

Mark1:
End Sub

Private Sub Command3_Click()
Text1.Text = ""
Text2.Text = ""
Text4.Text = ""
Text5.Text = ""
End Sub

Private Sub Form_Load()
Label2.Caption = number
Adodc1.RecordSource = "基本信息表"
Adodc1.Refresh
End Sub
