VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form4 
   Caption         =   "办公室办公界面"
   ClientHeight    =   5640
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10245
   LinkTopic       =   "Form4"
   ScaleHeight     =   5640
   ScaleWidth      =   10245
   StartUpPosition =   3  '窗口缺省
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   5880
      Top             =   120
      Visible         =   0   'False
      Width           =   1440
      _ExtentX        =   2540
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
      Connect         =   $"Form4.frx":0000
      OLEDBString     =   $"Form4.frx":0089
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc2"
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
   Begin VB.CommandButton Command11 
      Caption         =   "工资管理"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      TabIndex        =   34
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Frame Frame2 
      Caption         =   "查询操作"
      Height          =   2295
      Left            =   240
      TabIndex        =   33
      Top             =   3120
      Width           =   9975
      Begin VB.Frame Frame3 
         Caption         =   "综合查询"
         Height          =   1935
         Left            =   3480
         TabIndex        =   41
         Top             =   240
         Width           =   6255
         Begin VB.CommandButton Command14 
            Caption         =   "确定"
            Height          =   375
            Left            =   600
            TabIndex        =   55
            Top             =   1320
            Width           =   2295
         End
         Begin VB.TextBox Text13 
            Height          =   270
            Left            =   2520
            TabIndex        =   54
            Top             =   840
            Width           =   615
         End
         Begin VB.ComboBox Combo7 
            Height          =   300
            ItemData        =   "Form4.frx":0112
            Left            =   960
            List            =   "Form4.frx":0122
            TabIndex        =   53
            Top             =   840
            Width           =   1335
         End
         Begin VB.ComboBox Combo6 
            Height          =   300
            ItemData        =   "Form4.frx":0146
            Left            =   4680
            List            =   "Form4.frx":015C
            TabIndex        =   51
            Text            =   "全部"
            Top             =   1200
            Width           =   1335
         End
         Begin VB.ComboBox Combo5 
            Height          =   300
            ItemData        =   "Form4.frx":018C
            Left            =   4680
            List            =   "Form4.frx":019F
            TabIndex        =   49
            Text            =   "全部"
            Top             =   720
            Width           =   1335
         End
         Begin VB.ComboBox Combo4 
            Height          =   300
            ItemData        =   "Form4.frx":01CB
            Left            =   4680
            List            =   "Form4.frx":01D8
            TabIndex        =   47
            Text            =   "男女不限"
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox Text8 
            Height          =   270
            Left            =   2520
            TabIndex        =   43
            Top             =   360
            Width           =   615
         End
         Begin VB.ComboBox Combo3 
            Height          =   300
            ItemData        =   "Form4.frx":01EE
            Left            =   960
            List            =   "Form4.frx":01FE
            TabIndex        =   42
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "工龄"
            Height          =   180
            Left            =   360
            TabIndex        =   52
            Top             =   960
            Width           =   360
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "学历"
            Height          =   180
            Left            =   3960
            TabIndex        =   50
            Top             =   1200
            Width           =   360
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "所属部门"
            Height          =   180
            Left            =   3600
            TabIndex        =   48
            Top             =   720
            Width           =   720
         End
         Begin VB.Label Label16 
            Caption         =   "性别"
            Height          =   255
            Left            =   3960
            TabIndex        =   46
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "年龄"
            Height          =   180
            Left            =   360
            TabIndex        =   45
            Top             =   360
            Width           =   360
         End
      End
      Begin VB.CommandButton Command13 
         Caption         =   "确定"
         Height          =   375
         Left            =   2040
         TabIndex        =   38
         Top             =   1680
         Width           =   855
      End
      Begin VB.ComboBox Combo2 
         Height          =   300
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   37
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton Command12 
         Caption         =   "确定"
         Height          =   375
         Left            =   2040
         TabIndex        =   36
         Top             =   720
         Width           =   855
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         ItemData        =   "Form4.frx":0222
         Left            =   120
         List            =   "Form4.frx":0224
         Sorted          =   -1  'True
         TabIndex        =   35
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label14 
         Caption         =   "Label14"
         Height          =   255
         Left            =   3600
         TabIndex        =   44
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "职工号"
         Height          =   180
         Left            =   120
         TabIndex        =   40
         Top             =   1320
         Width           =   540
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "姓名"
         Height          =   180
         Left            =   120
         TabIndex        =   39
         Top             =   360
         Width           =   360
      End
   End
   Begin VB.CommandButton Command10 
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   32
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Caption         =   "基本信息操作"
      Height          =   3015
      Left            =   7560
      TabIndex        =   22
      Top             =   0
      Width           =   2535
      Begin VB.CommandButton Command9 
         Caption         =   "更新"
         Height          =   375
         Left            =   1320
         TabIndex        =   31
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton Command8 
         Caption         =   "删除"
         Height          =   375
         Left            =   1320
         TabIndex        =   30
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton Command7 
         Caption         =   "修改"
         Height          =   375
         Left            =   1320
         TabIndex        =   29
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton Command6 
         Caption         =   "添加"
         Height          =   375
         Left            =   1320
         TabIndex        =   28
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command5 
         Caption         =   "最后一条"
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton Command4 
         Caption         =   "第一条"
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "下一条"
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "上一条"
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "查看全部信息"
         Height          =   495
         Left            =   120
         TabIndex        =   23
         Top             =   2160
         Width           =   2175
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   5400
      Top             =   2040
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
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
      Connect         =   $"Form4.frx":0226
      OLEDBString     =   $"Form4.frx":02AE
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "基本信息表"
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
   Begin VB.TextBox Text12 
      DataField       =   "住址"
      DataSource      =   "Adodc1"
      Height          =   615
      Left            =   5280
      TabIndex        =   21
      Text            =   "Text12"
      Top             =   480
      Width           =   2295
   End
   Begin VB.TextBox Text11 
      DataField       =   "学历"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3840
      TabIndex        =   20
      Text            =   "Text11"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox Text10 
      DataField       =   "当前工作项目"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3840
      TabIndex        =   8
      Text            =   "Text10"
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox Text9 
      DataField       =   "联系电话"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3840
      TabIndex        =   7
      Text            =   "Text9"
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox Text7 
      DataField       =   "工龄"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3840
      TabIndex        =   6
      Text            =   "Text7"
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox Text6 
      DataField       =   "所在部门"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Text            =   "Text6"
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      DataField       =   "职务"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Text            =   "Text5"
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      DataField       =   "性别"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Text            =   "Text4"
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      DataField       =   "年龄"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      DataField       =   "职工号"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      DataField       =   "姓名"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "住址"
      Height          =   180
      Left            =   5280
      TabIndex        =   19
      Top             =   240
      Width           =   360
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "学历"
      Height          =   180
      Left            =   2880
      TabIndex        =   18
      Top             =   1560
      Width           =   360
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "当前工作项目"
      Height          =   180
      Left            =   2520
      TabIndex        =   17
      Top             =   2760
      Width           =   1080
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "联系电话"
      Height          =   180
      Left            =   2640
      TabIndex        =   16
      Top             =   2160
      Width           =   720
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "工龄"
      Height          =   180
      Left            =   2880
      TabIndex        =   15
      Top             =   960
      Width           =   360
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "性别"
      Height          =   180
      Left            =   240
      TabIndex        =   14
      Top             =   2160
      Width           =   360
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "所属部门"
      Height          =   180
      Left            =   2640
      TabIndex        =   13
      Top             =   360
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "职务"
      Height          =   180
      Left            =   240
      TabIndex        =   12
      Top             =   2760
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "年龄"
      Height          =   180
      Left            =   240
      TabIndex        =   11
      Top             =   1560
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "职工号"
      Height          =   180
      Left            =   120
      TabIndex        =   10
      Top             =   960
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "姓名"
      Height          =   180
      Left            =   240
      TabIndex        =   9
      Top             =   360
      Width           =   360
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form4.Hide
Form8.Show
End Sub

Private Sub Command10_Click()
End
End Sub

Private Sub Command11_Click()
Form4.Hide
Form10.Show
End Sub

Private Sub Command12_Click()
Adodc2.RecordSource = "select * from 基本信息表 where 姓名='" & Combo1.Text & "'"
Adodc2.Refresh
Text1.Text = Adodc2.Recordset.Fields("姓名").Value
Text2.Text = Adodc2.Recordset.Fields("职工号").Value
Text3.Text = Adodc2.Recordset.Fields("年龄").Value
Text4.Text = Adodc2.Recordset.Fields("性别").Value
Text5.Text = Adodc2.Recordset.Fields("职务").Value
Text6.Text = Adodc2.Recordset.Fields("所在部门").Value
Text7.Text = Adodc2.Recordset.Fields("工龄").Value
Text9.Text = Adodc2.Recordset.Fields("联系电话").Value
Text10.Text = Adodc2.Recordset.Fields("当前工作项目").Value
Text12.Text = Adodc2.Recordset.Fields("住址").Value
Text11.Text = Adodc2.Recordset.Fields("学历").Value
End Sub

Private Sub Command13_Click()
Adodc2.RecordSource = "select * from 基本信息表 where 职工号='" & Combo2.Text & "'"
Adodc2.Refresh
Text1.Text = Adodc2.Recordset.Fields("姓名").Value
Text2.Text = Adodc2.Recordset.Fields("职工号").Value
Text3.Text = Adodc2.Recordset.Fields("年龄").Value
Text4.Text = Adodc2.Recordset.Fields("性别").Value
Text5.Text = Adodc2.Recordset.Fields("职务").Value
Text6.Text = Adodc2.Recordset.Fields("所在部门").Value
Text7.Text = Adodc2.Recordset.Fields("工龄").Value
Text9.Text = Adodc2.Recordset.Fields("联系电话").Value
Text10.Text = Adodc2.Recordset.Fields("当前工作项目").Value
Text12.Text = Adodc2.Recordset.Fields("住址").Value
Text11.Text = Adodc2.Recordset.Fields("学历").Value
End Sub

Private Sub Command14_Click()
Form4.Hide
Form9.Show
End Sub

Private Sub Command2_Click()
Adodc1.Recordset.MovePrevious   '记录移到前一个
If Adodc1.Recordset.BOF Then    '如果为空，则移动到最后一个
Adodc1.Recordset.MoveLast
End If
End Sub

Private Sub Command3_Click()
Adodc1.Recordset.MoveNext        '记录移到后一个
If Adodc1.Recordset.EOF Then     '如果前一个为空，则移动到第一个
Adodc1.Recordset.MoveFirst
End If
End Sub

Private Sub Command4_Click()
Adodc1.Recordset.MoveFirst     '移动到第一个记录
End Sub

Private Sub Command5_Click()
Adodc1.Recordset.MoveLast      '移动到最后一个记录
End Sub

Private Sub Command6_Click()
Adodc2.RecordSource = "基本信息表"
Adodc2.Refresh
Adodc2.Recordset.MoveLast
num = Adodc2.Recordset.Fields("职工号").Value + 1
Form4.Hide
Form22.Show
End Sub

Private Sub Command7_Click()
Text1.Locked = False
Text3.Locked = False
Text4.Locked = False
Text5.Locked = False
Text6.Locked = False
Text7.Locked = False
Text9.Locked = False
Text10.Locked = False
Text11.Locked = False
Text12.Locked = False
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Command7.Enabled = False
Command8.Enabled = False
End Sub

Private Sub Command8_Click()
Dim X As Integer          '记录删除，删除前进行提示
X = MsgBox("是否真的要删除此条记录？", vbYesNo, "提示：")
If X = 6 Then
Adodc1.Recordset.Delete
Adodc1.Recordset.MoveFirst
End If
End Sub

Private Sub Command9_Click()

Text1.Locked = True
Text2.Locked = True
Text3.Locked = True
Text4.Locked = True
Text5.Locked = True
Text6.Locked = True
Text7.Locked = True
Text9.Locked = True
Text10.Locked = True
Text11.Locked = True
Text12.Locked = True
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command7.Enabled = True
Command8.Enabled = True
Adodc1.Recordset.Update
Adodc1.Refresh

Combo1.Clear
Combo2.Clear

Dim name1, number1 As String
Adodc1.Recordset.MoveFirst
Do
name1 = Adodc1.Recordset.Fields("姓名").Value
number1 = Adodc1.Recordset.Fields("职工号").Value
Combo1.AddItem name1
Combo2.AddItem number1
Adodc1.Recordset.MoveNext
Loop While Adodc1.Recordset.EOF <> True
Adodc1.Recordset.MoveFirst

End Sub

Private Sub Form_Load()
Text1.Locked = True
Text2.Locked = True
Text3.Locked = True
Text4.Locked = True
Text5.Locked = True
Text6.Locked = True
Text7.Locked = True
Text9.Locked = True
Text10.Locked = True
Text11.Locked = True
Text12.Locked = True

Adodc2.RecordSource = "select 职工号,姓名 from 基本信息表"
Adodc2.Refresh

Dim name, number As String
Adodc2.Recordset.MoveFirst
Do
name = Adodc2.Recordset.Fields("姓名").Value
number = Adodc2.Recordset.Fields("职工号").Value
Combo1.AddItem name
Combo2.AddItem number
Adodc2.Recordset.MoveNext
Loop While Adodc2.Recordset.EOF <> True


Dim zhiwu As String
zhiwu = Form2.Label9.Caption

If zhiwu <> "主任" Then
    Command6.Enabled = False
    Command7.Enabled = False
    Command8.Enabled = False
    Command9.Enabled = False
    Command11.Enabled = False
End If

End Sub

