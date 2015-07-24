VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "学生信息管理系统"
   ClientHeight    =   6720
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   6720
   ScaleWidth      =   8820
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command14 
      Caption         =   "退出系统"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6360
      TabIndex        =   34
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      Caption         =   "查询方式"
      Height          =   6255
      Left            =   5880
      TabIndex        =   30
      Top             =   120
      Width           =   2775
      Begin VB.CommandButton Command13 
         Caption         =   "综合查询"
         Height          =   855
         Left            =   240
         TabIndex        =   33
         Top             =   2880
         Width           =   2295
      End
      Begin VB.CommandButton Command10 
         Caption         =   "课程基本信息"
         Height          =   855
         Left            =   240
         TabIndex        =   32
         Top             =   1680
         Width           =   2295
      End
      Begin VB.CommandButton Command9 
         Caption         =   "学生基本信息"
         Height          =   735
         Left            =   240
         TabIndex        =   31
         Top             =   600
         Width           =   2295
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "基本信息管理"
      Height          =   6255
      Left            =   4320
      TabIndex        =   17
      Top             =   120
      Width           =   1575
      Begin VB.CommandButton Command8 
         Caption         =   "修改"
         Height          =   495
         Left            =   120
         TabIndex        =   25
         Top             =   3960
         Width           =   1095
      End
      Begin VB.CommandButton Command7 
         Caption         =   "更新"
         Height          =   495
         Left            =   120
         TabIndex        =   24
         Top             =   4680
         Width           =   1095
      End
      Begin VB.CommandButton Command6 
         Caption         =   "删除"
         Height          =   495
         Left            =   120
         TabIndex        =   23
         Top             =   3240
         Width           =   1095
      End
      Begin VB.CommandButton Command5 
         Caption         =   "添加"
         Height          =   495
         Left            =   120
         TabIndex        =   22
         Top             =   2640
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "最后一个"
         Height          =   495
         Left            =   120
         TabIndex        =   21
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "第一个"
         Height          =   495
         Left            =   120
         TabIndex        =   20
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "下一个"
         Height          =   495
         Left            =   120
         TabIndex        =   19
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "上一个"
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   120
      Top             =   6360
      Visible         =   0   'False
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=学生信息表\Student.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=学生信息表\Student.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "学生信息表"
      Caption         =   ""
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
   Begin VB.Frame Frame1 
      Caption         =   "基本信息"
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.TextBox Text13 
         DataField       =   "电话"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   29
         Text            =   "Text13"
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton Command11 
         Caption         =   "添加/修改照片"
         Enabled         =   0   'False
         Height          =   495
         Left            =   2640
         TabIndex        =   27
         Top             =   2280
         Width           =   1455
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   240
         Top             =   5040
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox Text12 
         DataField       =   "图片"
         DataSource      =   "Adodc1"
         Height          =   270
         Left            =   120
         TabIndex        =   26
         Text            =   "Text12"
         Top             =   5640
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.TextBox Text7 
         DataField       =   "个人语录"
         DataSource      =   "Adodc1"
         Height          =   2175
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   16
         Text            =   "Form1.frx":123EC
         Top             =   3360
         Width           =   1815
      End
      Begin VB.TextBox Text6 
         DataField       =   "年龄"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "Text6"
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox Text5 
         DataField       =   "家乡"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "Text5"
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         DataField       =   "专业"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   7.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "Text4"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         DataField       =   "学号"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "Text3"
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         DataField       =   "性别"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "Text2"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         DataField       =   "姓名"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   240
         Width           =   1095
      End
      Begin VB.PictureBox Picture1 
         Height          =   2895
         Left            =   1920
         Picture         =   "Form1.frx":123F2
         ScaleHeight     =   2835
         ScaleWidth      =   2115
         TabIndex        =   9
         Top             =   2760
         Width           =   2175
      End
      Begin VB.Label Label12 
         Caption         =   "电话"
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
         Left            =   120
         TabIndex        =   28
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "个人语录"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   8
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "照片"
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
         Left            =   1920
         TabIndex        =   7
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label Label6 
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
         Height          =   495
         Left            =   2040
         TabIndex        =   6
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "家乡"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "专业"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         TabIndex        =   4
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "学号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label2 
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
         Height          =   495
         Left            =   2160
         TabIndex        =   2
         Top             =   360
         Width           =   975
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
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Adodc1.Recordset.MovePrevious   '记录移到前一个
If Adodc1.Recordset.BOF Then    '如果为空，则移动到最后一个
Adodc1.Recordset.MoveLast
End If
End Sub

Private Sub Command10_Click()
Unload Form1                     '关闭基本窗口，打开课程相关查询窗口
Form3.Show
End Sub

Private Sub Command11_Click()
CommonDialog1.ShowOpen             '照片的添加和修改
Text12.Text = CommonDialog1.FileName    '将路径传给一个text控件
End Sub

Private Sub Command13_Click()
Unload Form1                       '关闭基本窗口，打开综合查询窗口
Form4.Show
End Sub

Private Sub Command14_Click()     '退出系统
End
End Sub

Private Sub Command2_Click()
Adodc1.Recordset.MoveNext        '记录移到前一个
If Adodc1.Recordset.EOF Then     '如果前一个为空，则移动到第一个
Adodc1.Recordset.MoveFirst
End If
End Sub

Private Sub Command3_Click()
Adodc1.Recordset.MoveFirst     '移动到第一个记录
End Sub

Private Sub Command4_Click()
Adodc1.Recordset.MoveLast      '移动到最后一个记录
End Sub

Private Sub Command5_Click()
Text1.Locked = False      '添加记录，此时将文本框解锁使得值能改变
Text2.Locked = False
Text3.Locked = False
Text4.Locked = False
Text5.Locked = False
Text6.Locked = False
Text7.Locked = False
Text13.Locked = False
Command11.Enabled = True
If Text3.Text = "" Then   '学号为主码，不能为空
MsgBox "学号不能为空！"
Exit Sub
End If
Adodc1.Recordset.AddNew
End Sub

Private Sub Command6_Click()
Dim X As Integer          '记录删除，删除前进行提示
X = MsgBox("是否真的要删除此条记录？", vbYesNo, "提示：")
If X = 6 Then
Adodc1.Recordset.Delete
Adodc1.Recordset.MoveFirst
End If
End Sub

Private Sub Command7_Click()
Text1.Locked = True
Text2.Locked = True   '进行更新操作，是改变后的窗口恢复如初
Text3.Locked = True
Text4.Locked = True
Text5.Locked = True
Text6.Locked = True
Text7.Locked = True
Text13.Locked = True
Command11.Enabled = False
Adodc1.Recordset.Update
End Sub

Private Sub Command8_Click()
Text1.Locked = False      '进行修改操作
Text2.Locked = False
Text3.Locked = False
Text4.Locked = False
Text5.Locked = False
Text6.Locked = False
Text7.Locked = False
Text13.Locked = False
Command11.Enabled = True   '使得照片能够修改
End Sub

Private Sub Command9_Click()
Unload Form1            '关闭当前窗口，打开有关学生信息查询的窗口
Form2.Show
End Sub




Private Sub Text12_Change()
Picture1.Picture = LoadPicture(Text12.Text)     '照片地址的载入
End Sub
