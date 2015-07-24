VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form3 
   Caption         =   "课程相关信息"
   ClientHeight    =   4830
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9465
   LinkTopic       =   "Form3"
   ScaleHeight     =   4830
   ScaleWidth      =   9465
   StartUpPosition =   3  '窗口缺省
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form3.frx":0000
      Height          =   3735
      Left            =   2760
      TabIndex        =   1
      Top             =   240
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   6588
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   360
      Top             =   4320
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      RecordSource    =   "select * from 选课表"
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
   Begin VB.Frame Frame1 
      Caption         =   "选择项"
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
      Begin VB.CommandButton Command2 
         Caption         =   "返回"
         Height          =   495
         Left            =   240
         TabIndex        =   7
         Top             =   3120
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "查询"
         Height          =   495
         Left            =   1440
         TabIndex        =   6
         Top             =   3120
         Width           =   855
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         Index           =   1
         ItemData        =   "Form3.frx":0015
         Left            =   480
         List            =   "Form3.frx":0064
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   1680
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         Index           =   0
         ItemData        =   "Form3.frx":0172
         Left            =   480
         List            =   "Form3.frx":01A9
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   960
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "课程名称"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "上课教师"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Sql As String
Public X As Integer
Dim db As Database
Private Sub Command1_Click()
On Error GoTo LABEL
Set db = OpenDatabase("学生信息表\Student.mdb")
Sql = "select 学生课程表.学号,成绩,课程名称,授课老师,学分 into Temp from 选课表,学生课程表 where 选课表.课程代号=学生课程表.课程代号"
db.Execute (Sql)
db.Close
Select Case X
Case 0        '按授课老师查询相关同学的信息
Sql = "select 学生信息表.姓名,学生信息表.学号,课程名称,成绩,授课老师,学分 from Temp,学生信息表 where 学生信息表.学号=Temp.学号 AND 授课老师='" & Combo1(0).Text & "'"
Case 1        '按课程名称查询相关同学的信息
Sql = "select 学生信息表.姓名,学生信息表.学号,课程名称,成绩,授课老师,学分 from Temp,学生信息表 where 学生信息表.学号=Temp.学号 AND 课程名称='" & Combo1(1).Text & "'"
End Select
Adodc1.RecordSource = Sql
Adodc1.Refresh
DataGrid1.Refresh
LABEL:
Set db = OpenDatabase("学生信息表\Student.mdb")
Sql = "DROP Table temp"
db.Execute (Sql)
db.Close
End Sub

Private Sub Command2_Click()
Unload Form3       '关闭当前窗口，返回主窗口
Form1.Show
End Sub

Private Sub Option1_Click(Index As Integer)
X = Index
End Sub
